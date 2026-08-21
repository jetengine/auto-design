# 双端开发指南（Mac 编辑 → Windows 运行）

> 本项目的运行时**必须**是 Windows + 已安装的 CATIA V5。Mac 只做代码编辑与提交。

---

## 0. 拓扑

```
┌──────────────┐    git push / pull    ┌────────────────────────┐
│  Mac (开发)  │  ──────────────────▶ │  Windows 工作站 (运行)  │
│  VS Code     │  ◀────────────────── │  CATIA V5 + Python      │
└──────────────┘                       └────────────────────────┘
```

推荐同步方式（任选其一）：

- **Git 仓库**（推荐）：Mac 推、Windows 拉。清晰、可回溯。
- **共享文件夹 / SMB**：省事，但不能追责，出问题难查。
- **VS Code Remote-SSH**：Mac 上开窗口，实际在 Windows 上编辑。适合联调阶段。

---

## 1. Windows 端一次性准备

在装了 CATIA V5 的那台 Windows 机器上：

1. **Python 3.10+**
   - 推荐 [python.org](https://www.python.org/downloads/windows/) 官方安装包；勾选 `Add python.exe to PATH`。
   - 注意：CATIA 是 64 位，Python 也要装 **64 位**，否则 COM 位数不匹配。

2. **拉代码**

   ```powershell
   git clone <你的仓库地址> C:\dev\AIchat
   cd C:\dev\AIchat
   ```

3. **建虚拟环境 + 装依赖**

   ```powershell
   py -3.10 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -e .
   ```

   这一步会装 `pywin32`。如果失败，通常是 pip 太老：`python -m pip install --upgrade pip`。

4. **权限对齐**（重要，容易踩坑）
   - **CATIA 和 Python 必须以相同权限运行**。
   - 若 CATIA 是普通用户身份启动，Python 也用普通用户 PowerShell 启动。
   - 若 CATIA 以管理员身份启动，Python 也要管理员 PowerShell —— COM 的 ROT 不跨完整性等级，跨了会 `GetActiveObject` 找不到对象。

---

## 2. Hello World —— 跑通 COM 通路

1. **手动启动 CATIA V5**（等它完全加载完，能看到主界面）。
2. 在同一台 Windows 的 PowerShell 里：
   ```powershell
   .\.venv\Scripts\Activate.ps1
   python scripts\hello_catia.py
   ```
3. 期望输出：
   ```
   ✅ 已连接到 CATIA。
     版本         : V5-6R2021
     Release / SP : 30 / 4
     文档数量     : 0
     活动文档     : （无）
     主窗口标题   : CATIA V5
   ```

只要这一步过了，后面的 MCP / Agent / 工具集都是**加法**，通路不会再出问题。

---

## 3. 常见故障速查

| 现象                                          | 大概率原因                     | 处理                                                                  |
| --------------------------------------------- | ------------------------------ | --------------------------------------------------------------------- |
| `GetActiveObject` → `Operation unavailable`   | CATIA 没启动，或权限不对齐     | 先启动 CATIA；对齐 Python 与 CATIA 的启动身份                         |
| `ImportError: No module named 'win32com'`     | pywin32 没装 / 装到别的解释器  | 确认已在虚拟环境里 `pip install -e .`                                 |
| 输出乱码                                      | Windows 控制台代码页不是 UTF-8 | `chcp 65001` 或在 PowerShell 里跑                                     |
| 一切正常但 `SystemConfiguration.Version` 报错 | CATIA 版本较老，属性名不同     | 换成 `app.SystemService.GetEnviron("VERSION")` 之类，先记下来后续适配 |
| `pywin32` 装不上                              | Python 是 32 位或版本 <3.10    | 换 64 位 Python 3.10+                                                 |

---

## 4. Mac 端能做什么

- 编辑代码（本文件、`src/catia_mcp/*.py`、`scripts/*.py`）。
- 提交、code review、写文档。
- **无法运行任何 CATIA 相关脚本** —— `import catia_mcp.catia_client` 在 Mac 上会直接 `RuntimeError`，这是刻意的。

如果想在 Mac 上跑单元测试，后续会引入一个 `mock_catia` 层，把 COM 对象替换成假对象。**Hello World 阶段不做**，避免过早抽象。

---

## 5. MCP Server —— 第一个工具 `get_catia_session`

Hello World 通过后，我们把 `session_info()` 包成了 MCP 工具。架构：

```
MCP client (Claude Desktop / VS Code)
        │  stdio + JSON-RPC
        ▼
   catia-mcp server (本项目)
        │  submit 到单 STA 线程
        ▼
   ComWorker（独占 STA 线程 + 唯一 CATIA 连接）
        │  COM
        ▼
   CATIA V5
```

> **为什么要 ComWorker**：CATIA COM 对象是单元线程（STA）绑定的，MCP 框架的回调可能在别的线程跑。所有 CATIA 调用都排队到这一个 STA 线程串行执行 —— 这就是"单 STA 串行 COM"的落地。53 个工具将来全部复用它。

### 5.1 安装 / 更新依赖（Windows）

```powershell
cd C:\Users\CNB0K20060\Desktop\AutoDesign\auto-design   # 你的实际路径
.\.venv\Scripts\Activate.ps1
pip install -e .    # 现在会一并装上 mcp
```

### 5.2 手动冒烟测试（不接 AI，先确认 server 能跑）

先启动 CATIA，再：

```powershell
python -c "from catia_mcp.com_worker import ComWorker; w=ComWorker(); w.start(); print(w.call(lambda c: c.session_info())); w.stop()"
```

能打印出 `CatiaSessionInfo(...)` 就说明 STA worker + COM 链路 OK。

### 5.3 挂到 MCP 客户端

**Claude Desktop**：编辑 `%APPDATA%\Claude\claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "catia": {
      "command": "C:\\Users\\CNB0K20060\\Desktop\\AutoDesign\\auto-design\\.venv\\Scripts\\catia-mcp.exe"
    }
  }
}
```

**VS Code**（`.vscode/mcp.json` 或用户设置）：

```json
{
  "servers": {
    "catia": {
      "type": "stdio",
      "command": "C:\\Users\\CNB0K20060\\Desktop\\AutoDesign\\auto-design\\.venv\\Scripts\\catia-mcp.exe"
    }
  }
}
```

> 路径换成你机器上 `catia-mcp.exe` 的实际位置（`pip install -e .` 后在 venv 的 `Scripts` 目录里）。

启动 CATIA → 重启 MCP 客户端 → 问它 **"CATIA 现在是什么版本、打开了哪个文档？"**，它应当调用 `get_catia_session` 并如实回答。

## 6. 第一个写操作 —— `create_box`（建模 + 体积回读验证）

第一个会**改变模型**的工具。它建一个长方体（草图矩形 + Pad 原生特征），并遵循"命令返回 ≠ 几何合格"原则：建模后 `Update` + 回读体积 + 与理论值比对，把证据一并返回。

### 6.1 无 AI 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\create_box_smoke.py
```

期望输出（关键字段）：

```
✅ 长方体建模完成，证据如下：
  document_name         : Part2.CATPart
  body_name             : PartBody
  sketch_name           : Sketch.1
  pad_name              : Pad.1
  update_ok             : True
  expected_volume_mm3   : 120000.0
  measured_volume_mm3   : 120000.0000...
  volume_match          : True
  relative_error        : ~0.0

判定： ✅ 几何合格
```

- `update_ok=True` → 特征树更新无红叉。
- `volume_match=True` → 回读体积与 L×W×H 吻合（容差 0.1%）。
- 两者都为真才算"几何合格"——这是第一次把"检查证据"落进代码。

> 若 `measured_volume_mm3` 为 `null`：说明 SPAWorkbench 测量在你这个环境不可用（不影响建模），把报错贴来，我换一种测量方式。

### 6.2 让 AI 建模

重载 MCP 客户端后，对它说：**"帮我建一个 100×60×20 的长方体。"**

它应当调用 `create_box(length_mm=100, width_mm=60, height_mm=20)`，然后根据返回的 `update_ok` / `volume_match` 告诉你是否成功。这就是主执行链的第一次完整跑通：

> 自然语言需求 → 工具调用 → CATIA 原生特征 → 检查证据 →（若不合格）迭代修复

## 7. 安全保存 + STEP 导出回读 —— `export_step_and_verify`（验证原则 5）

把活动 Part 安全保存为 `.CATPart`（可编辑特征树）并导出 `.STEP`（中性格式），再**重新打开 STEP 回读体积**与原始比对，防止导出过程几何丢失。

### 7.1 无 AI 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\export_step_smoke.py
```

它会一次跑完「建模 → 保存 CATPart → 导出 STEP → 回读验证」，输出到项目下 `out\`。关键字段：

```
② 导出/回读证据：
    catpart_saved         : True
    step_written          : True
    step_size_bytes       : 12345
    source_volume_mm3     : 120000.0
    reimported_volume_mm3 : 120000.0...
    volume_match          : True
    relative_error        : ~0.0

判定： ✅ 交付合格（建模 + STEP 回读均通过）
```

- `step_written=True` → STEP 文件真实落盘（实时文件验证）。
- `volume_match=True` → STEP 回读体积与原始吻合（容差 1%）。

> 若 `reimported_volume_mm3` 为 `null`：说明 STEP 在你环境打开成了 Product 或结构不同，回读测量没取到值（不影响导出本身）。把情况贴来，我补上 Product 遍历测量。

### 7.2 让 AI 交付

对 AI 说：**"把刚才的长方体导出成 STEP 并验证。"**
它应调用 `export_step_and_verify(step_path=...)`，根据 `step_written` / `volume_match` 回报是否交付合格。

至此，工程交付物齐备：**可编辑 `.CATPart` + 中性 `.STEP` + 验证证据**。

### 7.3 再下一步

1. 加 `measure_*` 只读测量族，扩充证据类工具。
2. 引入超时熔断 + 会话心跳看门狗（1、3 号风险）。
3. 扩到 Pocket / Fillet 等更多 Part Design 特征。
4. STEP 回读支持 Product 结构（多体/装配）。

## 8. 端到端：AI 端 → CATIA 全链路

前面的冒烟测试都是**进程内直接调 `ComWorker`**，绕过了 MCP 协议。这一节走**真实 MCP 协议**，与 Claude / VS Code Copilot 完全相同的路径：

```
MCP client ── stdio/JSON-RPC ──▶ catia_mcp.server ──▶ ComWorker(STA) ──▶ CATIA
```

### 8.1 可复现端到端测试（不依赖人在环）

在 **Windows**、CATIA 已启动、已 `pip install -e .`：

```powershell
python scripts\mcp_client_smoke.py
```

脚本以子进程拉起 server，完成 `initialize` → 列工具 → 依次调 `get_catia_session` / `create_box` / `export_step_and_verify`，并打印每步证据。关键判定：

```
✅ MCP 协议握手成功（initialize 完成）
✅ 发现工具：['get_catia_session', 'create_box', 'export_step_and_verify']
...
判定： ✅ 端到端链路跑通（MCP client → server → ComWorker → CATIA）
```

这一步只是把「LLM 决策调哪个工具」换成脚本写死的顺序，从而拿到确定性证据。

### 8.2 接真正的 AI（VS Code Copilot / Claude）

仓库已带 [.vscode/mcp.json](.vscode/mcp.json)（用 `${workspaceFolder}` 定位 venv 里的 `catia-mcp.exe`）。在 **Windows 上用 VS Code 打开本仓库**：

1. 确保已 `pip install -e .`（venv 里生成了 `catia-mcp.exe`）；
2. 启动 CATIA；
3. 在 Copilot Chat 的 Agent 模式里启用 `catia` MCP server；
4. 对 AI 说：**“看看 CATIA 现在什么状态，然后建个 100×60×20 的长方体并导出验证。”**

AI 应依次调用 `get_catia_session` → `create_box` → `export_step_and_verify`，并根据返回证据回报是否交付合格。这就是完整价值主张的端到端演示。

## 9. 第二种特征：挖槽 —— `add_pocket`（去料 + 体积差验证）

`create_box` 是加料，`add_pocket` 是**去料**——证明框架能做双向材料操作并验证。机制与建长方体同源、同样稳健（只用**偏移平面 + 草图 + Pocket**，不做脆弱的面/边拾取）：

```
XY 面上方 at_height 处建偏移平面 → 画矩形 → Pocket 向下挖 depth → Update → 体积差比对
```

### 9.1 无 AI 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\add_pocket_smoke.py
```

它建长方体 100×60×20，再在顶面居中挖 40×30、深 10 的盲槽。关键字段：

```
② 挖槽证据：
    pocket_name           : Pocket.1
    update_ok             : True
    volume_before_mm3     : 120000.0
    volume_after_mm3      : 108000.0
    expected_removed_mm3  : 12000.0
    measured_removed_mm3  : 12000.0
    volume_match          : True
    relative_error        : ~0.0

判定： ✅ 去料合格（建模 + 挖槽体积差均通过）
```

- `volume_match=True` → 实测去料（before−after）与理论 `pl×pw×depth` 吻合（容差 0.1%）。
- 约束：`depth < at_height`（只做盲槽，避免挖穿）；槽中心用长方体的 `L/2、W/2` 对齐。

### 9.2 让 AI 挖槽

对 AI 说：**“在刚才的长方体顶面正中挖一个 40×30、深 10 的槽。”**
它应调用 `add_pocket(pocket_length_mm=40, pocket_width_mm=30, depth_mm=10, at_height_mm=20, center_x_mm=50, center_y_mm=30)`，并根据 `update_ok` / `volume_match` 回报是否合格。

### 9.3 再下一步

1. Fillet（倒圆角）：需要边/面引用，比平面+草图更脆弱，单独一轮按证据推进。
2. Hole（孔）：草图点 + Hole 特征。
3. IGES 回读体积：导入曲面 `CloseSurface` 封成固体再测。
