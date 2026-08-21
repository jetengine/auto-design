"""导出能力探针 —— 定位 ExportData 失败到底是「许可证整体缺失」还是「STEP 单独没授权」。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/probe_export_formats.py

它会：
    1. 新建一个长方体 100×60×20（Sketch + Pad）；
    2. 对该 Part 逐个尝试 stp / step / igs / stl / wrl / model / cgr 导出；
    3. 打印每种格式的成功/失败与真实错误。

判读：
    - 若 stp/step/igs/stl 全失败 → CATIA 这台机器的转换器/许可证整体缺失。
    - 若仅 stp/step 失败，stl/igs 成功 → STEP 单独未授权，改用其它中性格式交付。
"""

from __future__ import annotations

import os
import sys
import time
import traceback


def main() -> int:
    try:
        from catia_mcp.com_worker import ComWorker
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "out"))
    tag = time.strftime("%Y%m%d_%H%M%S")

    worker = ComWorker()
    try:
        worker.start()

        worker.call(
            lambda c: c.create_box(100.0, 60.0, 20.0, part_name=f"ProbeBox_{tag}"),
            timeout=120.0,
        )

        results = worker.call(
            lambda c: c.probe_export_formats(out_dir),
            timeout=180.0,
        )

        print("导出格式探针结果：")
        any_ok = False
        for fmt, info in results.items():
            flag = "✅" if info["ok"] else "❌"
            any_ok = any_ok or info["ok"]
            print(f"    {flag} {fmt:<6} ok={info['ok']!s:<5} file={info['file']}")
            if not info["ok"]:
                print(f"           error: {info['error']}")

        print("\n判读：")
        step_ok = results.get("stp", {}).get("ok") or results.get("step", {}).get("ok")
        neutral_ok = any(results.get(f, {}).get("ok") for f in ("igs", "stl", "wrl", "cgr"))
        if step_ok:
            print("    ✅ STEP 可用 —— 之前失败多半是文件名/状态问题，回到 export_step_and_verify 即可。")
        elif neutral_ok:
            print("    ⚠️ STEP 单独不可用，但其它中性格式可导出 —— STEP 转换器未授权，改用可用格式交付。")
        elif any_ok:
            print("    ⚠️ 仅部分原生格式可用 —— 中性格式转换器整体缺失。")
        else:
            print("    ❌ 全部失败 —— 该 CATIA 安装无任何导出转换器，或权限/许可证问题。")
    except Exception:  # noqa: BLE001
        print("❌ 失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
