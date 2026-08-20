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

from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]

from .catia_client import CatiaClient, CatiaSessionInfo
from .com_worker import ComWorker

mcp = FastMCP("catia-mcp")

# 单一 STA worker，进程内唯一
_worker = ComWorker()


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

    info = _worker.call(_job, timeout=30.0)
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
    result = _worker.call(_job, timeout=120.0)
    return asdict(result)


@mcp.tool()
def export_step_and_verify(
    step_path: str,
    catpart_path: str | None = None,
) -> dict:
    """把当前活动 Part 安全保存为 CATPart（可选）并导出 STEP，再回读 STEP 验证体积。

    对应工程验证原则 5「STEP 回读」：导出后重新导入 STEP，比对体积，防止几何丢失。

    参数：
        step_path:    STEP 导出路径（.stp / .step，绝对路径）
        catpart_path: 可选，同时安全保存 .CATPart（.CATPart，绝对路径）

    返回（既是结果也是证据）：
        catpart_saved:          CATPart 是否成功落盘
        step_written:           STEP 文件是否真实存在且非空（实时文件验证）
        step_size_bytes:        STEP 文件大小
        source_volume_mm3:      导出前原实体体积
        reimported_volume_mm3:  STEP 回读后重新测得的体积
        volume_match:           回读体积是否与原始吻合（容差 1% 内）
        relative_error:         相对误差

    使用建议：step_written 与 volume_match 都为 true 才算交付合格。
    """

    def _job(client: CatiaClient):
        return client.export_step_and_verify(
            step_path=step_path,
            catpart_path=catpart_path,
        )

    # 保存 + 导出 + 重新打开测量，耗时较长
    result = _worker.call(_job, timeout=180.0)
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
