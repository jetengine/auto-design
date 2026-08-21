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
