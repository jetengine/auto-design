"""修饰特征 API 探针 —— 不再猜签名，直接问 CATIA 的类型库。

运行方式（在 **Windows**、CATIA 已启动，且当前有一个 Part 文档）：

    python scripts/probe_draft_api.py

── 为什么需要它 ──
首轮真机跑修饰特征，两个失败是同一个病因：**拿二手文档当一手事实**。

    AddNewSolidEdgeChamfer     这台机器上根本不存在
    AddNewChamfer(5 个参数)    "Invalid number of parameters."
    Part.CreateReferences      getattr 直接是 None

这些都不是"写错了"，是资料和这台机器对不上。而**类型库就在这台机器里**，
`ITypeInfo` 能把真实存在的方法名和参数表一条条读出来 —— 问它就好了，别猜。

和 probe_export_formats 一样，这是诊断工具：没有固定的合格标准，
所以**不进回归**。回归只放"能明确判对错"的东西。
"""

from __future__ import annotations

import sys
import traceback


def main() -> int:
    try:
        from catia_mcp.com_worker import ComWorker
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    worker = ComWorker()
    try:
        worker.start()
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 连接 CATIA 失败：{exc}", file=sys.stderr)
        return 2

    try:
        # 先建个小方块，保证有活动 Part 且 ShapeFactory 可用
        worker.call(
            lambda c: c.create_box(length_mm=20.0, width_mm=20.0, height_mm=20.0,
                                   part_name="ProbeDemo"),
            timeout=120.0, label="create_box(probe)",
        )
        report = worker.call(lambda c: c.probe_shape_api(), timeout=120.0,
                             label="probe_shape_api")
    except Exception:  # noqa: BLE001
        print("❌ 探针失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    for key, value in report.items():
        print(f"\n{key}:")
        if isinstance(value, list):
            if not value:
                print("    （空）")
            for item in value:
                print(f"    {item}")
        else:
            print(f"    {value}")

    print(
        "\n说明：把上面 shapefactory_draft / part_create_methods / references_* "
        "四段原样贴回来，拔模的签名就能一次钉死。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
