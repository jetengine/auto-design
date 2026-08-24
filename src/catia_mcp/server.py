"""CATIA MCP Server —— 第一个工具：get_catia_session。

架构：
    MCP client (Claude / VS Code)  ── stdio/JSON-RPC ──▶  本 server
                                                              │
                                                              ▼
                                                     ComWorker（单 STA 线程）
                                                              │
                                                              ▼
                                                       CATIA V5 (COM)

现在只暴露一个只读探测工具，验证「AI 能通过 MCP 问到 CATIA 真实状态」。
后续的 Sketch / Pad / Measure ... 都按同样的模式：定义 schema → submit 到 ComWorker。
"""

from __future__ import annotations

import sys
from dataclasses import asdict

if sys.platform != "win32":
    raise RuntimeError("catia_mcp.server 只能在 Windows 上运行（需要 CATIA COM）。")

# FastMCP 在不同发行里位置不同：官方 SDK 打包在 mcp.server.fastmcp，
# 也存在独立的 fastmcp 包。两处都试，给出可操作的报错。
try:
    from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
except ModuleNotFoundError:
    try:
        from fastmcp import FastMCP  # type: ignore[import-not-found]
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "找不到 FastMCP。你的 mcp 包里没有 mcp.server.fastmcp，也没有独立的 fastmcp。\n"
            "请安装带 FastMCP 的官方 SDK：\n"
            "    python -m pip install -U \"mcp>=1.2.0\"\n"
            "若仍失败，先查看已装版本：\n"
            "    python -c \"import mcp,os;print(mcp.__version__);import mcp.server as s;print(os.listdir(os.path.dirname(s.__file__)))\""
        ) from exc


from .catia_client import CatiaClient, CatiaSessionInfo
from .com_worker import ComWorker
mcp = FastMCP("catia-mcp")

# 单一 STA worker，进程内唯一
_worker = ComWorker()


@mcp.tool()
def catia_health() -> dict:
    """检查 AI↔CATIA 链路是否健康。任何建模失败/超时之后，**先调这个**。

    它分两层，且第一层不碰 CATIA —— 所以链路卡死时它仍然能回答：
        link.*        本地链路状态（线程、队列、当前卡在哪个任务、卡了多久）
        ping_ok / ping_error / catia_caption：一次廉价的真实 COM 往返

    关键字段：
        link.blocked:           true = 链路被一个卡死的任务堵住，后续调用会直接快速失败
        link.blocked_by:        堵住链路的任务名
        link.current_job_age_s: 当前任务已运行秒数
        ping_ok:                true = CATIA 真实可达

    处置建议：
        blocked=true 或 ping_ok=false 时，先提醒用户**去 CATIA 窗口关掉模态对话框**
        （最常见原因），关掉后仍不恢复再调 reconnect_catia。
    """
    link = _worker.health()
    result: dict = {"link": link}

    if link["blocked"]:
        # 链路已堵，再去 ping 只会白白多等一个超时
        result.update(
            ping_ok=False,
            ping_error=f"链路被「{link['blocked_by']}」堵住，已跳过 ping。",
            catia_caption=None,
        )
        return result

    try:
        caption = _worker.call(
            lambda c: c.ping(), timeout=10.0, label="catia_health.ping"
        )
        result.update(ping_ok=True, ping_error=None, catia_caption=caption)
    except Exception as exc:  # noqa: BLE001 —— 健康检查本身绝不能把异常抛给 AI
        result.update(ping_ok=False, ping_error=str(exc), catia_caption=None)
    return result


@mcp.tool()
def reconnect_catia() -> dict:
    """重建到 CATIA 的连接（不重启本进程）。仅在 catia_health 显示链路异常时使用。

    作用：丢弃卡死的 COM 线程，另起一个干净线程重新连 CATIA，使链路恢复可用。

    前提：若 CATIA 前台还开着模态对话框，**先关掉它再调本工具**，
    否则新连接一干活还会被同一个框挡住。

    返回：
        before:    重建前的链路快照（含堵死链路的任务名，用于事后定位）
        after:     重建后的链路快照
        recovered: 重建后一次真实 ping 是否成功
    """
    before = _worker.restart()
    try:
        caption = _worker.call(lambda c: c.ping(), timeout=15.0, label="reconnect.ping")
        recovered, ping_error = True, None
    except Exception as exc:  # noqa: BLE001
        caption, recovered, ping_error = None, False, str(exc)
    return {
        "before": before,
        "after": _worker.health(),
        "recovered": recovered,
        "ping_error": ping_error,
        "catia_caption": caption,
    }


@mcp.tool()
def get_catia_session() -> dict:
    """获取当前 CATIA 会话的只读状态快照。

    返回字段：
        system_configuration: CATIA 主版本标识
        release_number:       Release 号（如 34）
        service_pack:         Service Pack 号
        active_document_name: 当前活动文档名（无则为 null）
        document_count:       已打开文档数量
        caption:              CATIA 主窗口标题（供人眼二次核对）

    这是「能力与版本探测」的最小实现，也是所有后续建模操作前的健康检查入口。
    """

    def _job(client: CatiaClient) -> CatiaSessionInfo:
        return client.session_info()

    info = _worker.call(_job, timeout=30.0, label="get_catia_session")
    return asdict(info)


@mcp.tool()
def inspect_document(document_name: str | None = None) -> dict:
    """列出某个 Part 文档里有哪些 Body、每个 Body 里有哪些特征（只读）。

    这是测量的入口：`measure_body` 需要 body 名字，而你无从猜起，先用它看清结构。

    参数：
        document_name: 文档名（如 "Part3.CATPart"）；不传则用当前活动文档

    返回：
        document_name:        实际检查的文档
        open_documents:       当前打开的全部文档名（传错名字时照着改）
        is_part:              是否是零件文档；false 时（Product/Drawing）无法按零件测量
        bodies:               [{name, shape_count, shapes: [特征名...]}]
        body_count:           Body 数量
        active_document_name: 当前活动文档

    只做属性遍历，不会改动 CATIA 的任何状态。
    """

    def _job(client: CatiaClient):
        return client.inspect_document(document_name=document_name)

    result = _worker.call(_job, timeout=60.0, label="inspect_document")
    return asdict(result)


@mcp.tool()
def measure_body(
    document_name: str | None = None,
    body_name: str | None = None,
) -> dict:
    """测量一个已存在实体的体积、表面积、重心（只读，不改模型）。

    与其它工具的根本区别：它**不要求这个模型是你造的**。手工建的、别人发来的、
    上一轮造的，都能测 —— 这是唯一一个能对「既有模型」下结论的能力。

    参数：
        document_name: 文档名；不传则用当前活动文档
        body_name:     Body 名（如 "PartBody"）；不传则用第一个

    返回：
        volume_mm3:   体积（mm³）
        area_mm2:     表面积（mm²）
        cog_mm:       重心坐标 [x, y, z]（mm）
        cog_strategy: 重心是靠哪种 COM 调用姿势读到的（诊断用）
        cog_attempts: 成功前试过但不行的姿势（预期内，跨机器排障用）
        errors:       某一项真的测不出来时的原始报错（其余项照常返回）

    为什么三项都给：体积单独一项验证力有限 —— 拉伸方向反了、长宽写反了，
    体积可以一模一样，但**重心和表面积会露馅**。三项同时对上才算把几何钉死。
    """

    def _job(client: CatiaClient):
        return client.measure_body(document_name=document_name, body_name=body_name)

    result = _worker.call(_job, timeout=60.0, label="measure_body")
    return asdict(result)


@mcp.tool()
def create_box(
    length_mm: float,
    width_mm: float,
    height_mm: float,
    part_name: str | None = None,
) -> dict:
    """新建一个 Part，用原生特征（草图矩形 + Pad）建一个长方体，并回读体积验证。

    这是第一个写操作，产出保留可编辑特征树（Sketch + Pad）。

    参数：
        length_mm:  长（毫米，>0）
        width_mm:   宽（毫米，>0）
        height_mm:  高 / 拉伸厚度（毫米，>0）
        part_name:  可选，Part 文档命名

    返回（既是结果也是证据）：
        document_name / body_name / sketch_name / pad_name: 生成的特征树节点名
        update_ok:              特征树 Update 是否成功（无红叉）
        expected_volume_mm3:    理论体积 L×W×H
        measured_volume_mm3:    从 CATIA 回读的真实体积（测量失败为 null）
        volume_match:           回读体积是否与理论值吻合（容差 0.1% 内）
        relative_error:         相对误差

    使用建议：拿到返回后先看 update_ok 与 volume_match，二者都为 true 才算几何合格。
    """

    def _job(client: CatiaClient):
        return client.create_box(
            length_mm=length_mm,
            width_mm=width_mm,
            height_mm=height_mm,
            part_name=part_name,
        )

    # 建模比只读探测慢，给更长超时（超时熔断 = 安全边界）
    result = _worker.call(_job, timeout=120.0, label="create_box")
    return asdict(result)


@mcp.tool()
def add_pocket(
    pocket_length_mm: float,
    pocket_width_mm: float,
    depth_mm: float,
    at_height_mm: float,
    center_x_mm: float | None = None,
    center_y_mm: float | None = None,
) -> dict:
    """在当前活动 Part 的顶面挖一个矩形盲槽（Pocket 去料特征），并用体积差验证。

    这是第二种特征类型（去料），机制与 create_box 同源、同样稳健：
    偏移平面 + 草图矩形 + Pocket 向下挖，产出保留可编辑特征树。

    参数：
        pocket_length_mm/pocket_width_mm: 槽的长宽（毫米，>0）
        depth_mm:      挖深（毫米，须 < at_height_mm，只做盲槽避免挖穿）
        at_height_mm:  顶面所在 Z 高度（= 目标长方体高度）
        center_x_mm/center_y_mm: 槽中心（默认与长方体中心对齐；配合 create_box 用 L/2、W/2）

    返回（既是结果也是证据）：
        pocket_name:            生成的 Pocket 特征名
        update_ok:              特征树 Update 是否成功
        volume_before_mm3:      挖槽前体积
        volume_after_mm3:       挖槽后体积
        expected_removed_mm3:   理论去料体积 = pl×pw×depth
        measured_removed_mm3:   实测去料体积 = before − after
        volume_match:           实测去料是否与理论吻合（容差 0.1% 内）
        relative_error:         相对误差

    使用建议：update_ok 与 volume_match 都为 true 才算去料合格。
    先用 create_box 建体，再用同样的 length/width（中心 L/2、W/2）挖槽。
    """

    def _job(client: CatiaClient):
        return client.add_pocket(
            pocket_length_mm=pocket_length_mm,
            pocket_width_mm=pocket_width_mm,
            depth_mm=depth_mm,
            at_height_mm=at_height_mm,
            center_x_mm=center_x_mm,
            center_y_mm=center_y_mm,
        )

    result = _worker.call(_job, timeout=120.0, label="add_pocket")
    return asdict(result)


@mcp.tool()
def add_fillet(
    radius_mm: float,
    box_length_mm: float | None = None,
    box_width_mm: float | None = None,
    box_height_mm: float | None = None,
    propagation: str = "tangency",
) -> dict:
    """给当前活动 Part 的实体加恒定半径倒圆角（EdgeFillet），并用体积差验证。

    这是第三种特征类型（修饰特征），也是第一次涉及**边/面拾取**。
    实现刻意绕开逐条边的 BRep 引用（那种引用绑定具体拓扑、极易失效），
    改为把整个特征/实体交给 CATIA 展开成它的全部边 —— 因此这是**整体倒角**，
    不能挑单条边。

    参数：
        radius_mm:    圆角半径（毫米，>0，直径须小于最小边长，否则几何自相交）
        box_*_mm:     可选但**强烈建议传**。给出被倒角长方体的长宽高后，
                      系统能算出理论去料体积做硬验证；不传则只有弱验证。
        propagation:  "tangency"（默认，沿相切面传播）或 "minimal"

    返回（既是结果也是证据）：
        fillet_name:            生成的圆角特征名
        strategy:               实际生效的拾取策略（自证拾到的是什么）
        update_ok:              特征树 Update 是否成功
        volume_before/after_mm3: 倒角前后体积
        measured_removed_mm3:   实测去料 = before − after
        expected_removed_mm3:   理论去料（传了尺寸才有）
        volume_match:           实测与理论是否吻合（容差 0.1% 内）
        target_errors:          各失败拾取策略的原始报错（诊断用）

    使用建议：先 create_box，再用同样的长宽高调本工具，
    update_ok 与 volume_match 都为 true 才算倒角合格。
    """

    def _job(client: CatiaClient):
        return client.add_fillet(
            radius_mm=radius_mm,
            box_length_mm=box_length_mm,
            box_width_mm=box_width_mm,
            box_height_mm=box_height_mm,
            propagation=propagation,
        )

    result = _worker.call(_job, timeout=120.0, label="add_fillet")
    return asdict(result)


@mcp.tool()
def add_chamfer(
    length_mm: float,
    angle_deg: float = 45.0,
    box_length_mm: float | None = None,
    box_width_mm: float | None = None,
    box_height_mm: float | None = None,
    propagation: str = "tangency",
) -> dict:
    """给当前活动 Part 的实体全部边倒斜角（Chamfer），并用体积差验证。

    与 add_fillet 同一条边拾取路径，换把刀而已 —— 直棱代替圆弧。
    什么时候用它而不是倒圆角：去毛刺、装配导入角、焊接坡口这类场景要的是斜面；
    要减应力集中或做外观过渡才用圆角。

    参数：
        length_mm:   斜角第一条边长（毫米，>0）
        angle_deg:   斜角角度，默认 45°。**只有 45° 有精确解验证**，
                     其余角度只做弱验证（会如实把 expected_removed_mm3 报成 null，
                     不会拿近似值假装验过）。
        box_*_mm:    可选但强烈建议传，传了才有硬验证。
        propagation: "tangency"（默认）或 "minimal"

    返回（既是结果也是证据）：
        strategy:               实际生效的方法与枚举值（自证）
        objects_chamfered:      实际倒了多少条边（长方体应为 12）
        expected/measured_removed_mm3、volume_match、relative_error

    使用建议：update_ok 与 volume_match 都为 true 才算合格。
    """

    def _job(client: CatiaClient):
        return client.add_chamfer(
            length_mm=length_mm,
            angle_deg=angle_deg,
            box_length_mm=box_length_mm,
            box_width_mm=box_width_mm,
            box_height_mm=box_height_mm,
            propagation=propagation,
        )

    result = _worker.call(_job, timeout=120.0, label="add_chamfer")
    return asdict(result)


@mcp.tool()
def add_shell(
    thickness_mm: float,
    box_length_mm: float | None = None,
    box_width_mm: float | None = None,
    box_height_mm: float | None = None,
) -> dict:
    """把当前活动 Part 的实体抽成薄壳（自动去掉顶面开口），并用体积差验证。

    这是第一个**必须指名道姓挑一个面**的特征。挑面不按索引（顺序不保证），
    而是逐面测重心、取 Z 最大的那个 —— 换任何模型都成立。
    被选中面的面积和重心会一并返回，所以「它到底挑了哪个面」有据可查。

    参数：
        thickness_mm: 壁厚（毫米，>0，向内偏移，外轮廓不变）
        box_*_mm:     可选。传了才算得出理论去料体积 = (L−2t)(W−2t)(H−t)。

    返回（既是结果也是证据）：
        removed_face:   被去掉面的 index / area_mm2 / cog_mm（自证挑对了没有）
        face_candidates: 一共拾到几个面（长方体应为 6）
        expected/measured_removed_mm3、volume_match、relative_error

    使用建议：壁厚要满足 2t < min(长,宽) 且 t < 高，否则抽不出内腔（会直接报错拦下）。
    """

    def _job(client: CatiaClient):
        return client.add_shell(
            thickness_mm=thickness_mm,
            box_length_mm=box_length_mm,
            box_width_mm=box_width_mm,
            box_height_mm=box_height_mm,
        )

    result = _worker.call(_job, timeout=180.0, label="add_shell")
    return asdict(result)


@mcp.tool()
def add_draft(
    angle_deg: float,
    box_length_mm: float | None = None,
    box_width_mm: float | None = None,
    box_height_mm: float | None = None,
) -> dict:
    """给当前活动 Part 的四个侧面加拔模斜度（底面为中性面），并用体积差验证。

    铸造/注塑件脱模必需的特征。三样东西同时给对才成立：被拔模的侧面、
    中性面（保持不变的基准，这里取底面）、拔模方向（Z 轴）。

    参数：
        angle_deg: 拔模角（度，0 < a < 45，常见 1~5）
        box_*_mm:  可选。传了才算得出理论体积变化。

    返回（既是结果也是证据）：
        faces_drafted:          实际拔了几个面（长方体应为 4）
        neutral_face:           中性面的面积与重心（自证选的是底面）
        measured_delta_mm3:     实测体积变化，**带符号**
        expected_outward_mm3 / expected_inward_mm3:
                                上大下小 / 上小下大 两种情形的精确解
        matched_direction:      实测符合哪一个（两者大小不等，所以这仍是硬验证）
        volume_match、relative_error

    使用建议：先看 matched_direction 是否有值 —— 为 null 说明两种情形都对不上，
    那就是真出问题了，去看 strategy 和 target_errors。
    """

    def _job(client: CatiaClient):
        return client.add_draft(
            angle_deg=angle_deg,
            box_length_mm=box_length_mm,
            box_width_mm=box_width_mm,
            box_height_mm=box_height_mm,
        )

    result = _worker.call(_job, timeout=180.0, label="add_draft")
    return asdict(result)


@mcp.tool()
def create_box_family(
    variants: list[dict],
    output_dir: str | None = None,
) -> dict:
    """按参数表一次生成一族长方体（可选倒角），每个变体各自验证体积。

    这是把前面所有单件能力**放大到规模**的工具：做 20 个变体和做 1 个，
    对你来说心智负担一样，但对人来说是 20 次重复劳动 + 20 次出错机会。

    参数：
        variants: 变体列表，每项：
            length_mm / width_mm / height_mm : 必填，正数
            fillet_radius_mm                 : 可选，需满足 2r < 最短边
            name                             : 可选，用作 Part 名与文件名
            一次最多 50 个（批量期间独占 CATIA 链路，太长会让其它调用一直排队）
        output_dir: 给了就每个变体**存盘后关闭**，会话保持干净；
                    不给就全部留在 CATIA 里（注意 20 个变体 = 20 个文档）

    返回：
        requested / succeeded / failed: 请求数、成功数、失败数
        all_verified:                   是否全部建成且体积对上（一眼定论）
        elapsed_s:                      总耗时
        documents_left_open:            没存盘、留在会话里的文档数
        variants:                       每个变体一条明细，**失败的也有**，
                                        含 ok / relative_error / error

    读法：先看 all_verified；为 false 时再到 variants 里找 ok=false 的那几条，
    每条都带自己的失败原因，不用重跑整批去定位。
    """

    def _job(client: CatiaClient):
        return client.create_box_family(variants=variants, output_dir=output_dir)

    # 超时按变体数给，而不是拍一个固定值：20 个变体本来就该比 1 个等得久，
    # 用固定值要么冤枉大批次，要么给小批次留了没意义的长窗口。
    budget = min(60.0 + 20.0 * len(variants or []), 900.0)
    result = _worker.call(_job, timeout=budget, label=f"create_box_family(n={len(variants or [])})")
    return asdict(result)



@mcp.tool()
def export_step_and_verify(
    step_path: str,
    catpart_path: str | None = None,
) -> dict:
    """把当前活动 Part 安全保存为 CATPart（可选）并导出中性格式，再回读验证体积。

    对应工程验证原则 5「回读」：导出后重新导入，比对体积，防止几何丢失。
    格式策略：默认优先 STEP；若该 CATIA 未授权 STEP 转换器，自动降级为 IGES。

    参数：
        step_path:    首选导出路径（.stp / .step / .igs，绝对路径）；降级时自动换扩展名
        catpart_path: 可选，同时安全保存 .CATPart（.CATPart，绝对路径）

    返回（既是结果也是证据）：
        catpart_saved:          CATPart 是否成功落盘（核心交付物）
        step_written:           中性格式文件是否真实存在且非空（实时文件验证）
        step_size_bytes:        导出文件大小
        format_used:            实际成功使用的格式（如 stp / igs）
        source_volume_mm3:      导出前原实体体积
        reimported_volume_mm3:  回读后重新测得的体积（曲面型导入可能为 null）
        volume_match:           回读体积是否与原始吻合（容差 1% 内）
        relative_error:         相对误差
        export_error:           导出失败时的真实 COM 错误（诊断用）

    使用建议：catpart_saved 与 step_written 为 true 即交付闭环成立；
    volume_match 为体积回读加分项（中性格式受导入类型影响，可能为 null）。
    """

    def _job(client: CatiaClient):
        return client.export_step_and_verify(
            step_path=step_path,
            catpart_path=catpart_path,
        )

    # 保存 + 导出 + 重新打开测量，耗时较长
    result = _worker.call(_job, timeout=180.0, label="export_step_and_verify")
    return asdict(result)


def main() -> None:
    """stdio 模式启动 MCP server。"""
    _worker.start()  # 先把 STA 线程 + CATIA 连接拉起，连不上直接报错退出
    try:
        mcp.run(transport="stdio")
    finally:
        _worker.stop()


if __name__ == "__main__":
    main()
