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
3. 扩到更多 Part Design 特征（Pocket 已完成，见第 9 节；下一步 Fillet / Hole）。
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

`create_box` 是加料，`add_pocket` 是**去料**——证明框架能做双向材料操作并验证。机制与建长方体同源（只用**平面 + 草图 + Pocket**，不做脆弱的面/边拾取）：

```
XY 面上方 at_height 处建偏移平面 → 画矩形 → Pocket 向下挖 depth → Update → 体积差比对
```

### 9.1 踩到的坑：偏移平面挂不上草图

先后两次都挂在同一行：

```
sketch = body.Sketches.Add(offset_plane)
pywintypes.com_error: (-2147352567, ..., (0, 'CATIASketches', 'The method Add failed', ...))
```

真因不是参数包没包 `CreateReferenceFromObject`（两种写法报同一个错），而是那个平面**根本没有几何**：`AddNewPlaneOffset` 造出的平面必须先被挂进某个容器并 `Update()` 才被真正算出来；而原先用的 `body.InsertHybridShape(plane)` **只在 Part 开了 Hybrid Design 时才成立**，这台机器上不成立。

修正：改用**几何图形集**（`part.HybridBodies.Add()`）——每个 Part 都有这个容器，不依赖任何开关。`_make_pocket_sketch()` 把它做成降级链：

| 策略          | 做法                                                                                           | 挖料方向 |
| ------------- | ---------------------------------------------------------------------------------------------- | -------- |
| **A**         | `HybridBodies.Add()` → `AppendHybridShape` → `InWorkObject=hb` → `Update` → 在偏移平面上建草图 | 顶面向下 |
| **B**（兜底） | 直接用 `PlaneXY`（`create_box` 已证可用），把 Pocket 方向翻成 +Z                               | 底面向上 |

结果里带 `strategy` / `pocket_from` / `plane_errors` 三个**自证字段**，不用猜走了哪条路。

### 9.2 无 AI 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\add_pocket_smoke.py
```

它建长方体 100×60×20，再在顶面居中挖 40×30、深 10 的盲槽。**实际输出（已跑通）**：

```
② 挖槽证据：
    pocket_name           : Pocket.1
    update_ok             : True
    volume_before_mm3     : 120000.0
    volume_after_mm3      : 108000.0
    expected_removed_mm3  : 12000.0
    measured_removed_mm3  : 12000.0
    volume_match          : True
    relative_error        : 0.0
    strategy              : hybrid_body_offset_plane(ref=False)
    pocket_from           : top
    plane_errors          : None

判定： ✅ 去料合格（建模 + 挖槽体积差均通过）
```

- `volume_match=True` → 实测去料（before−after）与理论 `pl×pw×depth` 完全吻合（容差 0.1%）。
- `strategy=hybrid_body_offset_plane(ref=False)` → 策略 A 生效，平面对象直接传即可，无需包 Reference。
- 约束：`depth < at_height`（只做盲槽，避免挖穿）；槽中心用长方体的 `L/2、W/2` 对齐。

### 9.3 让 AI 挖槽

对 AI 说：**“在刚才的长方体顶面正中挖一个 40×30、深 10 的槽。”**
它应调用 `add_pocket(pocket_length_mm=40, pocket_width_mm=30, depth_mm=10, at_height_mm=20, center_x_mm=50, center_y_mm=30)`，并根据 `update_ok` / `volume_match` 回报是否合格。

### 9.4 再下一步

1. Fillet（倒圆角）：需要边/面引用，比平面+草图更脆弱，单独一轮按证据推进。
2. Hole（孔）：草图点 + Hole 特征。
3. IGES 回读体积：导入曲面 `CloseSurface` 封成固体再测。

## 10. 健壮性：超时熔断 + 阻塞自诊断 + 重建恢复

### 10.1 要防的是什么

单 STA 线程保证了 COM 调用串行有序，但代价是 **一个调用卡死 = 整条链路卡死**。这不是理论风险：CATIA 只要在前台弹一个模态框（许可证提示、文件覆盖确认、错误对话框），当前 COM 调用就**永远不返回**，没有任何超时。

改造前的后果很难受：

- 卡住的那次调用超时失败；
- STA 线程仍死在里面，**后续每次调用都排队 → 各自等满超时 → 全部失败**；
- 而且报的都是「超时」这种没有信息量的错，AI 和人都不知道到底出了什么事。

### 10.2 四道防线

| 防线           | 位置                                                   | 作用                                                             |
| -------------- | ------------------------------------------------------ | ---------------------------------------------------------------- |
| **掐掉弹框源** | `CatiaClient.connect()` 设 `DisplayFileAlerts = False` | 让文件类模态框根本不弹，从源头减少死锁                           |
| **超时熔断**   | `ComWorker.call(..., timeout, label)`                  | 到点即失败，并把链路标记为「被 label 堵住」                      |
| **快速失败**   | 下一次 `call()` 的前置检查                             | 链路已知被堵时**立即** `CatiaBlockedError`，不再陪等一个完整超时 |
| **重建恢复**   | `ComWorker.restart()`                                  | 丢弃卡死线程，另起干净线程重连，无需重启进程                     |

两个设计要点值得记：

1. **`health()` 不碰 COM。** 健康检查如果自己也要走 COM，链路一卡它就跟着卡 —— 那就永远问不出「到底出了什么事」。所以它只读本地状态（线程存活、队列深度、当前任务名与已运行秒数），**卡死时也一定能返回**。真实 COM 往返（`ping`）是单独一层，且链路已堵时直接跳过。
2. **卡死线程是「丢弃」不是「杀死」。** 卡在 COM 调用里的线程无法被安全终止（强杀会破坏 COM 运行时状态）。它是 daemon 线程，会挂到 CATIA 那边的框被关掉后自行结束。我们只是不再等它，代价是短暂多一个僵尸线程 —— 这是清醒的取舍，不是遗漏。

### 10.3 两个新工具

- **`catia_health`** —— 任何失败后 AI 应该先调的。返回 `link.blocked` / `link.blocked_by` / `link.current_job_age_s` / `ping_ok`。
- **`reconnect_catia`** —— 仅在 health 显示异常时用。返回 `before`（含堵死链路的任务名，留证）/ `after` / `recovered`。

错误信息本身就是处置指引，AI 能直接转述给用户：

```
「create_box」超过 120s 未返回，当前卡在「create_box」已 133s。
CATIA 通常是被模态对话框挡住了（许可证/覆盖确认/错误框）。
处理：切到 CATIA 窗口关闭对话框；必要时调用 reconnect_catia 重建链路。
```

### 10.4 冒烟测试（会**人为制造一次卡死**）

```powershell
python scripts\health_smoke.py
```

它占住 STA 线程 25 秒来模拟 COM 卡死 —— 与真实卡死在「线程被占住、后续任务全堵在队列」这点上完全等价，但可控、可重复。**实际输出（10/10 通过）**：

```
① 正常态健康检查
  ✅ 链路空闲 — blocked=False, current_job=None, queue_depth=0
  ✅ CATIA 可达 — caption='CATIA V5'

② 人为制造卡死（占住 STA 线程 25s）
  ✅ 受害调用超时并给出可执行提示 — 耗时 3.0s，错误指名了阻塞源

③ 卡死态下的健康检查（关键：它自己不能也跟着卡住）
  ✅ health 立即返回 — 耗时 0.000s
  ✅ 如实报告 blocked — blocked_by='fake_hang'
  ✅ 报出卡了多久 — current_job_age_s=3.5

④ 熔断：后续调用应立即失败，而不是再等满一个超时
  ✅ 熔断生效（快速失败） — 耗时 0.000s（预算 60s，实际没等）

⑤ 恢复：重建链路
  ✅ 重建前快照留证 — blocked_by='fake_hang'
  ✅ 重建后链路可用 — caption='CATIA V5'
  ✅ 重建后不再阻塞 — blocked=False, jobs_done=2, restarts=1

判定： ✅ 健壮性合格（超时熔断 / 阻塞自诊断 / 重建恢复 全部生效）
```

第 ③④ 两条是重点：**健康检查自己不卡**（0.000s）+ **熔断不空等**（预算 60s，实际 0.000s），这两点决定了链路出事时系统还能不能说人话。第 ⑤ 条的 `restarts=1` 证明恢复走的是重建路径，不是碰巧自愈。

### 10.5 尚未覆盖

1. 卡死线程回收：僵尸线程要等 CATIA 那边解除阻塞才结束，多次 restart 会累积。
2. 无人值守：现在靠 AI/人看到 health 再决定重连，没有后台看门狗自动触发。
3. CATIA 进程崩溃/退出：目前表现为连接错误，尚未做自动重连退避。

## 11. 第三种特征：倒圆角 —— `add_fillet`（第一次碰「几何引用」）

### 11.1 它和前两个特征本质不同

|                            | 输入是什么                 | 脆点                               |
| -------------------------- | -------------------------- | ---------------------------------- |
| `create_box`（Pad）        | **参数**（长宽高）         | 基本没有                           |
| `add_pocket`（Pocket）     | **参数** + 一个基准面      | 平面得先有几何（第 9 节的坑）      |
| `add_fillet`（EdgeFillet） | **几何引用**（拾取哪些边） | 引用绑定具体拓扑，改尺寸就可能失效 |

CATIA 的边引用长这样：

```
REdge:(Edge:(Face:(Brp:(Pad.1;2);None:();Cf11:());Face:(Brp:(Pad.1;1);...
```

这串 BRep 名字写死了「第几个面、第几条边」。上游特征一改，编号漂移，引用失效 —— 这是所有 CAD 自动化最经典的脆点，也是选 Fillet 作为第三个特征的原因：**早点把它暴露出来**，因为后续所有需要选面/选边的特征（倒角、抽壳、拔模）都会踩同一个坑。

### 11.2 规避策略：不写死拓扑名字

本实现**不去构造逐条边的 BRep 名字**，而是把整个特征/实体交给 CATIA，让它自己展开成全部边：

| 候选拾取对象              | 说明                               |
| ------------------------- | ---------------------------------- |
| `body.Shapes.Item(count)` | 实体里最后一个特征（通常就是 Pad） |
| `body`                    | 整个实体                           |

每个候选再叉乘「包不包 `CreateReferenceFromObject`」×「`AddNewSolidEdgeFilletWithConstantRadius` / `AddNewEdgeFilletWithConstantRadius`」两种工厂方法 —— 因为方法名与可接受的对象类型在不同 CATIA 版本/配置下并不统一，而失败只在调用瞬间以 COM 错误暴露。**与其猜，不如把组合跑一遍并把每次失败原文记下来**（`target_errors`），生效的那条记进 `strategy`。

代价说清楚：只能**整体倒角**，不能挑单条边。挑边留到确有需求时再按证据推进。

### 11.3 验证用的是精确解，不是估算

倒圆角的去料量不好直接算（角料形状怪异）。所以反过来算**倒圆后的实体**，把它拆成四块互不重叠、每块都是初等体积的部分：

$$V_{\text{rounded}} = \underbrace{abc}_{\text{内芯}} + \underbrace{2r(ab+ac+bc)}_{\text{6 块面板}} + \underbrace{\pi r^2 (a+b+c)}_{\text{12 段 1/4 圆柱}} + \underbrace{\tfrac{4}{3}\pi r^3}_{\text{8 个 1/8 球 = 1 整球}}$$

其中 $a = L-2r,\; b = W-2r,\; c = H-2r$。去料量 $= LWH - V_{\text{rounded}}$，是**精确值**。

以 40×30×20、R5 为例：

| 项                     | 值               |
| ---------------------- | ---------------- |
| 内芯 30×20×10          | 6000.000         |
| 面板 2·5·(600+300+200) | 11000.000        |
| 圆柱 π·25·60           | 4712.389         |
| 球 4π·125/3            | 523.599          |
| **倒圆后体积**         | **22235.988**    |
| 原体积 40×30×20        | 24000            |
| **理论去料**           | **1764.012 mm³** |

体积能对上，说明**12 条边全被拾到且都倒成功了** —— 这是引用层面真正生效的硬证据，比「没报错」强得多。

### 11.4 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\add_fillet_smoke.py
```

期望：

```
② 倒圆角证据：
    fillet_name           : EdgeFillet.1
    strategy              : last_shape(Pad.1)/ref=True/AddNewSolidEdgeFilletWithConstantRadius
    update_ok             : True
    volume_before_mm3     : 24000.0
    volume_after_mm3      : 22235.99
    measured_removed_mm3  : 1764.01
    expected_removed_mm3  : 1764.012
    volume_match          : True
    target_errors         : None

判定： ✅ 倒圆角合格（边拾取生效，去料体积与精确解吻合）
```

**读结果的顺序**：先看 `volume_match`（对不对），再看 `strategy`（怎么做到的），失败时看 `target_errors`（哪几条路走不通、原文报错是什么）。

### 11.5 让 AI 倒角

对 AI 说：**“把这个 40×30×20 的块所有边倒 R5 圆角。”**
它应调用 `add_fillet(radius_mm=5, box_length_mm=40, box_width_mm=30, box_height_mm=20)`，并根据 `update_ok` / `volume_match` 回报是否合格。

> 传 `box_*_mm` 很关键：不传就只有「体积变小了」这种弱验证，传了才有理论值可比对。
