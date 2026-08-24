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

### 11.2 第一次尝试：把整个特征丢给 CATIA —— 被实测否决

最省事的想法是不碰单条边，直接把整个 Pad / Body 交给 CATIA，让它自己展开成全部边。于是把 8 种组合全跑了一遍（`{last_shape, body}` × `{包 Reference, 不包}` × `{AddNewSolidEdgeFillet…, AddNewEdgeFillet…}`）。**全部失败**，但失败方式讲清了原因：

| 组合        | 报错                               | 含义                                                  |
| ----------- | ---------------------------------- | ----------------------------------------------------- |
| `ref=False` | `Type mismatch`                    | 方法必须收 `Reference`，不收裸 COM 对象               |
| `ref=True`  | `The method AddNew…Fillet… failed` | 引用类型没问题，但**整个特征/实体不是合法的倒角对象** |

这就是把候选组合全跑一遍、并把每次失败原文记进 `target_errors` 的价值：**否定结论也是结论**。如果只试一种写法，得到的只有一句没有信息量的 "method failed"，根本判断不出该往哪个方向改。

结论：CATIA 要的是真正的**边引用**，没有捷径。

### 11.3 第二次尝试：让 CATIA 自己把边找出来

既然必须要边引用，那手写 BRep 名字行不行？不行 —— 那等于把「第几个面、第几条边」硬编码进代码，换个模型立刻失效，而且拼错了同样只会得到一句 "method failed"。

真正的解法是**用 CATIA 自己的选择集搜索**：

```python
selection.Search("Topology.CGMEdge,all")
refs = [selection.Item(i).Reference for i in range(1, count + 1)]
```

拿到的 `Reference` 与人在界面上点选那条边**完全等价**，由 CATIA 自己生成 —— 天然合法，且不含任何硬编码拓扑名字。然后：

```python
fillet = shape_factory.AddNewSolidEdgeFilletWithConstantRadius(refs[0], mode, r)
for ref in refs[1:]:
    fillet.AddObjectToFillet(ref)     # 其余边追加进同一个特征
```

特征树上就是干净的一个 `EdgeFillet`，而不是 12 个碎特征。

搜索串在不同版本/语言环境下写法略有差异，所以按序试了三种（`Topology.CGMEdge,all` / `,sel` / 带引号形式），生效的记进 `strategy`。实测第一种即通过。

#### 候选数 ≠ 实体边数

实测有个必须讲清的细节：**搜索是全文档范围的**。40×30×20 的长方体捞到 **16 个候选**——12 条实体边 + 4 条草图线。那 4 条被 CATIA 拒绝（`AddObjectToFillet failed`）是**预期行为，不是故障**。

所以结果里同时给两个数：

| 字段               | 含义                   |
| ------------------ | ---------------------- |
| `edge_candidates`  | 搜索到的候选总数（16） |
| `objects_filleted` | 实际倒角成功数（12）   |

差值有据可查，不把预期内的拒绝伪装成错误 —— 否则每次成功的倒角都会附带一堆吓人的红色报错，久而久之就没人看 `target_errors` 了。

### 11.4 验证用的是精确解，不是估算

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

### 11.5 冒烟测试（Windows，CATIA 已启动）

```powershell
python scripts\add_fillet_smoke.py
```

期望：

```
② 倒圆角证据：
    fillet_name           : EdgeFillet.1
    strategy              : Search(Topology.CGMEdge,all)/AddNewSolidEdgeFilletWithConstantRadius/edges=12of16
    update_ok             : True
    volume_before_mm3     : 24000.0
    volume_after_mm3      : 22235.987758454303
    measured_removed_mm3  : 1764.0122415456972
    expected_removed_mm3  : 1764.012244017009
    volume_match          : True
    relative_error        : 1.4e-09
    objects_filleted      : 12
    edge_candidates       : 16
    target_errors         : ['4/16 个候选被 CATIA 拒绝（通常是草图线等非实体边，属预期内）']

判定： ✅ 倒圆角合格（边拾取生效，去料体积与精确解吻合）
```

`relative_error` 是 **1.4e-09** —— 不是「差不多对」，是与精确解在浮点精度内一致。这说明 12 条边确实全被拾到、且都按 R5 倒成功了。

**读结果的顺序**：

1. `objects_filleted` —— 倒了几条边？长方体必须是 12。这一步失败说明**拾取**有问题。
2. `volume_match` —— 倒出来对不对？这一步失败说明**几何**有问题。
3. `strategy` —— 具体是哪条路走通的。
4. `target_errors` —— 失败时看哪几条路走不通、原文报错是什么。

把「拾取」和「几何」分成两个独立可观测的指标，出问题时不用猜是哪一层。

### 11.6 让 AI 倒角

对 AI 说：**“把这个 40×30×20 的块所有边倒 R5 圆角。”**
它应调用 `add_fillet(radius_mm=5, box_length_mm=40, box_width_mm=30, box_height_mm=20)`，并根据 `update_ok` / `volume_match` 回报是否合格。

> 传 `box_*_mm` 很关键：不传就只有「体积变小了」这种弱验证，传了才有理论值可比对。

---

## 12. 回归：把七个冒烟脚本串成一条命令

### 12.1 为什么现在必须做

到这里已经有 7 个冒烟脚本，每个都**单独**验证通过过。问题在于它们共用同一份 `catia_client.py` 和 `com_worker.py`。

做倒圆角时我顺手动了 `part.InWorkObject` —— Pocket 会不会被打断？没人知道，除非每次都全跑一遍。而**一条命令跑不完的回归，等于没有回归**：只要需要人手敲 7 次、记 7 个判定标准，实际执行率就是零。

```powershell
python scripts\regression.py              # 全跑
python scripts\regression.py --list       # 看有哪些步骤
python scripts\regression.py fillet pocket  # 只跑关键字匹配的
python scripts\regression.py --skip mcp   # 跳过最慢的端到端
```

### 12.2 顺序不是随便排的

先证明链路通，再证明能力对，最后跑最慢的端到端：

| 步骤      | 脚本                   | 超时 | 验证内容                        |
| --------- | ---------------------- | ---- | ------------------------------- |
| `connect` | `hello_catia.py`       | 60s  | COM 链路 / 版本 / 会话          |
| `health`  | `health_smoke.py`      | 120s | 超时熔断 + 阻塞自诊断 + 恢复    |
| `box`     | `create_box_smoke.py`  | 180s | 加料特征（Sketch + Pad）        |
| `pocket`  | `add_pocket_smoke.py`  | 180s | 去料特征（偏移平面 + Pocket）   |
| `fillet`  | `add_fillet_smoke.py`  | 180s | 几何引用（边拾取 + Fillet）     |
| `export`  | `export_step_smoke.py` | 240s | 交付（STEP，降级 IGES）         |
| `mcp`     | `mcp_client_smoke.py`  | 300s | MCP 端到端（起服务器 + 调工具） |

`probe_export_formats.py` **不在**回归里——它是探测当前机器支持哪些格式的诊断工具，没有固定通过标准，每台机器答案都不同。把它塞进回归只会制造假失败。

### 12.3 四个设计取舍

1. **子进程隔离**。每个脚本单独起进程。一个脚本把 CATIA 搞挂了，下一个还能跑，且能自证「我是在什么状态下失败的」。同进程串跑的话，第一个泄漏的 COM 引用会污染后面所有结论——那种回归的绿灯是骗人的。

2. **不吞输出**。子进程 stdout/stderr 直连终端，证据实时滚出来。捕获再回放会让 3 分钟的运行看起来像卡死了。

3. **硬超时**。CATIA 卡死时不给「再等等看」的余地——超时即杀，记 `TIMEOUT`。**回归脚本自己挂住是最糟的失败模式**：它本来是用来发现卡死的。

4. **前置失败即中止**。连不上 CATIA 时，后面 6 个脚本会吐出 6 份一模一样的连接错误，把真正的第一现场淹掉。所以 `connect` 失败直接停，其余标 `SKIP`。

### 12.4 判定符号

退出码与各冒烟脚本的约定一致：

| 符号 | 判定      | 含义                                      |
| ---- | --------- | ----------------------------------------- |
| ✅   | `PASS`    | 退出码 0                                  |
| ❌   | `FAIL`    | 退出码 1 —— 跑起来了，但证据不合格        |
| 🚫   | `ENV`     | 退出码 2 —— 环境问题（没装包/CATIA 没开） |
| ⏱    | `TIMEOUT` | 超时被杀                                  |
| ⏭   | `SKIP`    | 前置中止的连带结果，重跑它没有意义        |

分开 `FAIL` 和 `ENV` 是有用的：前者是**代码退步了**，后者是**机器没准备好**。混在一起会让人在环境问题上白查半天代码。

结尾只列真正跑过且没过的步骤供单独重跑，`SKIP` 不列。

### 12.5 基线（CATIA V5R34 SP3，已实测）

首次（7 步）**7/7 通过，合计 23.4s**；加测量族（8 步）**8/8，24.3s**；加批量建族（9 步）：

```
  ✅ PASS       0.3s  connect   COM 链路 / 版本 / 会话
  ✅ PASS       4.0s  health    超时熔断 + 阻塞自诊断 + 重建恢复
  ✅ PASS       1.8s  box       加料特征（Sketch + Pad）
  ✅ PASS       1.5s  pocket    去料特征（偏移平面 + Pocket）
  ✅ PASS       1.9s  fillet    几何引用（边拾取 + EdgeFillet）
  ✅ PASS       2.2s  measure   只读测量（体积/面积/重心 三独立证据）
  ✅ PASS      13.3s  family    批量建族（规模化 + 单个失败不拖垮整批）
  ✅ PASS       3.6s  export    交付（STEP，不可用则降级 IGES）
  ✅ PASS       6.2s  mcp       MCP 端到端（起服务器 + 调工具）

  9/9 通过，合计 34.8s
```

**35 秒**——这个数字比「9/9 通过」更重要。回归的执行率由它决定：半分钟的回归每次改动都会跑，5 分钟的只在提交前跑，20 分钟的没人跑。

从 24.3s 涨到 34.8s，**涨的 10.5s 有 10.4s 是 family 一个步骤贡献的**，其余八步合计反而略降。这是可接受的：批量是唯一一个「一步做六件事」的步骤，它的耗时天然是别人的六倍。但也因此，它是**唯一一个需要盯住耗时上限的步骤**——真要再涨，就把回归版变体数从 6 降到 3，批量语义用 3 个和 6 个证明力完全一样。

基线里几个值得记的实测点：

- `health` 那 4.0s 里有 3.0s 是**故意等的**（人为制造卡死后验证熔断），不是性能问题。
- `export` 稳定落到 `format_used: igs`、`reimported_volume_mm3: null` —— 这台机器 STEP 翻译器没许可，降级路径按预期生效。**降级是已知结论，不是失败**，所以判 PASS。
- `fillet` 稳定在 `edges=12of16`、`relative_error=1.4e-09`，与单独跑时一致 —— 串跑没有污染它。
- `measure` 三项与单跑完全一致，且它是唯一**只读**的步骤：跑再多遍也不改变任何模型。
- `family` 串跑 13.3s，单跑 11.5s，**慢了约 1.8s**：前面六步已经在 CATIA 里留下模型，文档变多之后每次 `Documents.Add` 和存盘都要贵一点。这个差值本身就是"串跑不等于单跑之和"的证据。

### 12.6 顺带证伪的一件事

跑之前我担心的是：做倒圆角时动过 `part.InWorkObject`，会不会把 Pocket 打断？

结果 `pocket` 步骤 `relative_error=0.0` 照常通过。这个担心被证伪了——**但价值不在于"没坏"，而在于以后每次改动都能用一条命令重新问一遍这个问题**，不必再靠记忆和运气。

---

## 13. 只读测量族 —— `inspect_document` / `measure_body`

### 13.1 它解决的是一个之前绕不过去的问题

到第 12 节为止，所有工具都有同一个隐含前提：**这个模型是我造的**。

`create_box` 知道理论体积是 L×W×H，因为 L/W/H 是它自己的入参。`add_fillet` 知道该去掉多少料，因为半径是它自己传的。一旦模型不是本轮造的——手工建的、同事发来的、上一轮 AI 造的——现有工具**一句话也说不出来**。

测量族是第一个不依赖"我造的"这个前提的能力。

### 13.2 为什么必须是三条证据，不能只有体积

前面每个写操作都用「体积吻合」自证。但体积单独一项的验证力是有限的：

| 错误                       | 体积         | 会被谁抓到           |
| -------------------------- | ------------ | -------------------- |
| Pad 拉反了方向             | **一模一样** | 重心（z 跑到另一侧） |
| 长宽写反（40×30 → 30×40）  | **一模一样** | 重心                 |
| 形状完全不同、体积恰好相同 | **一模一样** | 表面积               |
| 有几条边没倒到             | 偏小一点点   | 重心漂移（更灵敏）   |

所以 `measure_body` 一次返回**体积 + 表面积 + 重心**。三者互相独立，同时对上才算把几何真正钉死。这是**第二条独立证据轴**，不是把同一个数再读一遍。

### 13.3 两个工具，先看清再测

`measure_body` 需要一个 body 名字，而 AI 无从猜起。所以必须配一个 `inspect_document`：

```python
inspect_document(document_name=None)
    → open_documents / is_part / bodies:[{name, shape_count, shapes:[...]}]

measure_body(document_name=None, body_name=None)
    → volume_mm3 / area_mm2 / cog_mm / cog_strategy / errors
```

两个细节：

- **找不到时把"现在开着哪些"一并报出来**。光说「没找到 Part9.CATPart」等于让人重猜；把 `open_documents` 附上，AI 自己就能改对。
- **`is_part: false` 如实报告**。Product / Drawing / STEP 导入件都可能没有 `.Part`，此时诚实说"测不了"，比编一个 0 出来强。

### 13.4 出参数组：一个「不报错但说谎」的调用姿势

`GetCOG` 不是返回值，而是**出参数组**（`CATSafeArrayVariant`）。pywin32 传这种参数的几种写法，实测结果是：

| 姿势                                      | 结果                                             |
| ----------------------------------------- | ------------------------------------------------ |
| `VARIANT(VT_ARRAY\|VT_BYREF\|VT_VARIANT)` | ❌ `Objects for SAFEARRAYS must be sequences...` |
| `VARIANT(VT_ARRAY\|VT_BYREF\|VT_R8)`      | ❌ 同上                                          |
| 直接传 `list`                             | ⚠️ **不报错，但缓冲区一个字节没变**              |
| `SystemService.Evaluate` VBA 跳板         | ✅ **本机唯一走通的**                            |

第三种才是危险的那个：pywin32 把 list **按值**传了进去，CATIA 老老实实往那份副本里写了重心，我们读的原 list 纹丝不动，于是拿到一个漂亮的 `(0, 0, 0)`。

> **它没有失败，它在说谎。** 静默返回错误数据的策略，比抛异常的危险得多——后者会停下来，前者会把假数据一路带进结论。

#### 对策：哨兵值，而不是「见到零就当失败」

缓冲区不填 0，填一个真实重心绝无可能取到的量级：

```python
sentinel = -1.2345678901e9
```

调用完若三个分量**还是哨兵值**，就判定「出参没被写回」，换下一种姿势。

关键是不能图省事写成「结果是 (0,0,0) 就算失败」——**居中建模的零件重心本来就是原点**，那样会把正确结果误杀。哨兵区分的是「没被写」，而不是「值恰好为零」。

#### 真正跑通的：VBA 跳板

绕开 pywin32 的出参 marshalling，让 CATIA 内部接住数组再当**返回值**递出来：

```python
vba = ("Function GetCOGArray(m)\n"
       "    Dim c(2)\n"
       "    m.GetCOG c\n"
       "    GetCOGArray = c\n"
       "End Function")
app.SystemService.Evaluate(vba, lang, "GetCOGArray", [measurable])
```

语言枚举在不同版本文档里对不上号，所以按 `(2, 0, 1)` 依次试，成的记进 `cog_strategy`。本机是 `lang=2`。代价：需要脚本执行权限，企业环境可能被锁——所以它排在最后，前面三种能成就轮不到它。

### 13.5 单位：CATIA 自己就不一致

同一个 `Measurable` 对象上：

| 属性     | 单位   | 换算   |
| -------- | ------ | ------ |
| `Volume` | m³     | ×1e9   |
| `Area`   | m²     | ×1e6   |
| `GetCOG` | **mm** | **×1** |

这不是笔误，是实测：40×30×20 的块，`GetCOG` 直接给回 `(20, 15, 10)`。第一版按 SI 惯例乘了 1e3，得到 `(20000, 15000, 10000)`——差整整 1000 倍。

教训很直白：**别推断单位，逐个用已知精确解钉死**。也正因为差 3 个数量级，冒烟测试一跑就露馅，这个坑不可能被带过去。

### 13.6 尝试记录 ≠ 错误

成功路径上那三条失败的姿势，是这台机器的**预期行为**，不是故障。所以它们进 `cog_attempts`，不进 `errors`：

- `cog_attempts` —— 试过但不行的姿势，跨机器排障时是金子
- `errors` —— 真正导致某一项测不出来的原因

只有当重心彻底没拿到时，`cog_attempts` 才会一并进 `errors`（那时它们不再是背景噪声，而是唯一的线索）。

这和 §11 倒圆角那个「4 条草图线被拒」是同一个原则：**不把预期内的失败伪装成错误**，否则每次成功都附带一堆红字，久而久之就没人看了。

### 13.7 冒烟测试

```powershell
python scripts\measure_smoke.py
```

它建一个 40×30×20 的块，测三项；再倒 R5 圆角，**再测一次**。精确解：

|          | 体积 (mm³) | 表面积 (mm²) | 重心 (mm)             |
| -------- | ---------- | ------------ | --------------------- |
| 长方体   | 24000      | 5200         | (20, 15, 10)          |
| 倒 R5 后 | 22235.988  | 4399.115     | **(20, 15, 10) 不变** |

表面积的精确解同样是拆出来的（$a=L-2r$ 等）：

$$A_{\text{rounded}} = \underbrace{2(ab+ac+bc)}_{\text{6 块平面}} + \underbrace{2\pi r(a+b+c)}_{\text{12 段 1/4 圆柱侧面}} + \underbrace{4\pi r^2}_{\text{8 个 1/8 球面}}$$

实测（Part9.CATPart）：

```
长方体   volume 24000.0            area 5200.0            rel_err 0.00e+00 / 0.00e+00
倒 R5 后 volume 22235.987758454303 area 4399.114857511296 rel_err 1.11e-10 / 3.55e-13
重心     (20, 15, 10) → 倒角后 (19.999999999634, 14.99999999995, 10.00000000005)
```

**重心不变这一条是这个测试里最有意思的**：倒圆角对三个中心面都是对称的，所以重心必须原地不动，实测漂移量 **3.7e-10 mm**。一旦真的漂了，说明有边没倒到或倒错了边——而这恰恰是**体积几乎测不出来**的那类错误（少倒一条边，体积只差百分之几，重心却立刻偏）。

### 13.8 让 AI 测量

对 AI 说：**"看看现在打开的这个零件，体积和重心是多少？"**

它应先 `inspect_document` 看清结构，再 `measure_body`。这是第一次 AI 能对**它没参与建的模型**给出有依据的回答。

---

## 14. 批量建族 —— `create_box_family`（换的是量级，不是功能）

### 14.1 为什么这一步才是重点

前面九个工具，单次调用都已经可靠。但**「AI 帮我设计」和「AI 帮我点鼠标」的差别，恰恰在于能不能一次生成并验证 20 个变体**。

对人来说，做 20 遍是 20 次重复劳动 + 20 次出错机会，第 13 个上手滑了自己都不知道。对 AI 来说，做 20 遍和做 1 遍的心智负担一样——**前提是每一个都能自证**。前面所有关于「精确解验证」的坚持，价值到这里才完全兑现：批量的可信度不能靠抽查。

```python
create_box_family(variants=[...], output_dir=None)
```

### 14.2 两类失败，反应必须相反

这是批量语义里最容易写错的地方：

| 失败类型     | 例子                    | 正确反应                 |
| ------------ | ----------------------- | ------------------------ |
| **规格错**   | 负数尺寸、`2r ≥ 最短边` | **整批不动**，一个都别建 |
| **运行时错** | CATIA 里几何求解失败    | **就地记下，继续往下**   |

规格错在碰 CATIA 之前就能查出来。既然能提前知道，就绝不能建了 6 个再报错——那等于留下 6 份要手工收拾的垃圾，而且**批量的每一步都不可回滚**（建出来的文档不会自己消失）。

运行时错则相反：第 7 个挂了，不该让第 8~20 个白等。

所以校验全部前置、执行逐个 try。

### 14.3 失败的变体必须留一条记录

```python
len(result.variants) == result.requested   # 恒等式
```

悄悄跳过失败项，汇总里就会 20 变 19——**而没有人会发现少了哪个**。所以明细条数恒等于请求数，靠 `ok` 字段区分，每条失败都带自己的 `error`。

定位失败不需要重跑整批，这在一次几分钟的批量里很关键。

### 14.4 汇总才是产物

20 份原始返回没人看得完，AI 的上下文也会被淹没。所以先给结论：

```
requested / succeeded / failed / all_verified / elapsed_s / documents_left_open
```

读法：**先看 `all_verified`**；为 false 再去 `variants` 里找 `ok=false` 的那几条。

### 14.5 顺手解决了文档堆积

批量把一个之前"还不值得修"的问题变成了现实问题：**20 个变体 = 20 个文档**。前面几轮回归下来 CATIA 里已经积了 16 份，再来一次批量就没法看了。

给了 `output_dir` 就**存盘后关闭**，会话保持干净；不给就全留着。`documents_left_open` 如实报出会话污染量。

> 这正是"等它真的碍事了再做"的兑现——不是提前设计，是问题真的出现了才动手。

### 14.6 两条安全边界

**批次上限 50**。批量期间**独占单 STA 链路**，健康检查和其它调用全部排队。一个失手写成 5000 的请求会让链路瘫痪几十分钟——超时熔断也救不了，因为它确实在正常工作。

**文件名消毒**。`name` 可能来自 AI 生成，直接拼进路径的话 `name="../../x"` 就能写到 `output_dir` 之外。所以只保留字母数字和 `_-`，并再用 `commonpath` 复核一次结果确实落在 `output_dir` 里内。

### 14.7 冒烟测试：测的是**批量语义**，不是功能

```powershell
python scripts\family_smoke.py
```

前面每个冒烟测「这个功能对不对」，这个测「批量行为对不对」：

- 6 个变体（3 板 + 3 带倒角）全部存盘，逐个用精确解核对
- 汇总自洽性：`requested == 明细条数`、`documents_left_open == 0`
- 文件真实落盘（实时文件验证，不是"没报错就算写了"）
- **故意混入 2 个非法变体**，验证前置校验整批拒绝，且错误信息**指名是第几个**
- 51 个变体，验证上限保护

理论体积（供比对）：

| 变体    | 尺寸     | R   | 体积 (mm³) |
| ------- | -------- | --- | ---------- |
| Plate_A | 60×40×8  | —   | 19200      |
| Plate_B | 60×40×12 | —   | 28800      |
| Plate_C | 60×40×16 | —   | 38400      |
| Round_A | 50×40×30 | 3   | 59109.079  |
| Round_B | 50×40×30 | 5   | 57592.182  |
| Round_C | 50×40×30 | 8   | 54093.120  |

> **功能正确 ≠ 批量正确**。一个 `for` 循环里 `try` 写错位置，就能把"20 个全成功"变成"第 1 个成功、剩下 19 个被静默吞掉"——而单件测试一个都抓不到。

### 14.8 让 AI 建族

对 AI 说：**"给我一组 60×40 的板，厚度从 8 到 16 每 4 一档，都存到 out\family。"**

它应把这句话展开成 `variants` 列表调 `create_box_family`，然后**只需回报 `all_verified`**——因为每个变体都已经自己证明过自己了。

### 14.9 实测基线（已在 Windows 验证）

```
requested 6 / succeeded 6 / failed 0 / all_verified True
elapsed_s 11.546 / documents_left_open 0

✅ Plate_A  19200.0     rel 0.00e+00   边 None
✅ Plate_B  28800.0     rel 1.26e-16   边 None
✅ Plate_C  38400.0     rel 0.00e+00   边 None
✅ Round_A  59109.079   rel 4.13e-12   边 12
✅ Round_B  57592.182   rel 2.96e-11   边 12
✅ Round_C  54093.120   rel 1.61e-10   边 12
```

三个值得记的实测点：

- **11.5 秒 6 个变体，约 1.9s/件**——其中带倒角的比纯 Pad 贵。按这个速率，50 个变体的上限约合 100 秒，落在单次调用可接受的范围内，说明上限值定得合理。
- **`documents_left_open: 0`**，6 个 CATPart 全部落盘后关闭。存盘 + 关闭这条路径不弹对话框，`DisplayFileAlerts = False` 在 `SaveAs`/`Close` 上同样生效。
- **误差随倒角半径单调上涨**（4.13e-12 → 2.96e-11 → 1.61e-10），仍比容差 1e-3 低七个数量级。这是 NURBS 求交的正常表现：圆角越大，参与运算的曲面越多。

前置校验和上限保护也按预期动作，错误信息**指名第几个变体**：

```
✅ NegDim   : 第 1 个变体：长宽高必须为正数，收到 (-10.0, 20.0, 5.0)。
✅ FatFillet: 第 1 个变体：倒角半径 6.0 过大 —— 2r 必须小于最短边 10.0，…
✅ 上限     : 一次最多 50 个变体，收到 51 个。
```

**两条路径都没碰 CATIA 就返回了**——这正是前置校验的意义：非法批次的代价是零。

---

## 15. 修饰特征三兄弟 —— `add_chamfer` / `add_shell` / `add_draft`

### 15.1 这是抄作业，但抄作业也分三档

最难的一关早在倒圆角那一步就趟平了：**怎么拿到 CATIA 认的几何引用**。剩下的三个特征都建在那条路上。但它们的难度是递增的，而且**每一档新增的东西完全不一样**：

| 特征        | 复用什么       | 新增什么                             |
| ----------- | -------------- | ------------------------------------ |
| **Chamfer** | 整条边拾取路径 | 只有枚举值不确定                     |
| **Shell**   | 引用机制       | 第一次要**指名道姓挑某一个面**       |
| **Draft**   | 引用机制       | 要**一组面 + 一个中性面 + 一个方向** |

看清这个梯度很重要：如果以为三个都是"换个方法名"，那实际难度会在第二个上突然跳一档，而人还以为是自己写错了。

### 15.2 Chamfer：枚举猜错为什么不会静默通过

`CatChamferMode` 里「长度 + 角度」这一档，各处资料写作 `0` / `1` / `2` 的都有。押一个值然后在别的机器上炸掉，是最没意思的失败方式。

所以按序试，把**实际生效的值写进 `strategy`**：

```
Search(Topology.CGMEdge,all)/AddNewSolidEdgeChamfer(mode=1)/edges=12of16
```

关键在于：**猜错不会假装成功**。如果误用成「两个长度」模式，第二个参数 `45` 会被当成 45mm 的第二条边长——在 40×30×20 的体上直接吃穿，要么 CATIA 报错，要么体积对不上。

> 精确解验证在这里的角色变了：它不是"锦上添花的复核"，**它就是枚举的判据本身**。

45° 斜角的精确解和倒圆角同一套拆法，只是把圆的零件换成直的：

```
内芯 a·b·c + 6 块面板 2d(ab+ac+bc) + 12 段三棱柱 2d²(a+b+c) + 8 个角块 2d³
```

那 8 个角块合起来正好是棱长 d 的**菱形十二面体**（体积 `2d³`）。这不是巧合——把边长 `2d` 的正方体十二条边各切 `d`，剩下的就是它。40×30×20 切 `d=4`，理论去料 **2496.000 mm³**。

**只有 45° 有精确解**。其它角度的八个角块是不规则多面体，那就如实把 `expected_removed_mm3` 报成 `null`，**宁可不验，也不用近似值假装验过了**。

### 15.3 Shell：第一次必须挑面，而挑面不能按索引

倒角可以「把搜到的边全都要」。抽壳不行——它必须回答"**去掉哪一个面**"。

诱惑是按索引挑：`refs[2]` 就是顶面。这在自己这台机器上能跑，然后在换个模型或换个 R 版之后**悄悄挑错**——而且不报错，只是做出个别的东西。这类失败比崩溃危险得多，和之前 COG 返回 `(0,0,0)` 是同一类。

所以：**不按索引挑面，按测量挑面**。

逐个面读出重心，取重心 Z 最大的那个。这是几何事实，换任何模型都成立。代价是每个面一次 `GetMeasurable` 加一次重心读取（重心还得走那条 VBA 蹦床）——长方体六个面而已，这笔开销买的是"换模型不会错"。

顺带一提：这一步**白嫖了测量族**。`_read_cog_mm` 那三轮调试当时是为了测量工具做的，现在直接成了挑面的判据。能力互相垫脚，是因为每一层都做成了独立可用的东西，而不是塞进某个大函数里。

返回里带上被选中面的**面积和重心**，所以"它到底挑了哪个面"有据可查，不用去 CATIA 里肉眼找：

```
removed_face: index=?, area_mm2=1200.0, cog_mm=(20.0, 15.0, 20.0)
```

内腔精确解：四周各让出 `t`、底部让出 `t`、顶部开口不让，即 `(L−2t)(W−2t)(H−t)`。壁厚 3 时 = **13872.000 mm³**。

### 15.4 Draft：三样东西同时对，而且未必报错

拔模要同时给：被拔模的**一组侧面**、**中性面**（拔模时保持不变的基准，取底面）、**拔模方向**（Z 轴，用 PlaneXY 的法向表达）。

错一个，结果就不是想要的形状——**而且未必报错**，可能安静地做出个别的东西。

所以这里的验证比前面更狠。拔模到底是加料（上大下小）还是去料（上小下大），取决于角度符号和 CATIA 的方向约定，**这是实测才能定的事**。与其押一个约定，不如把两种情形的精确解都算出来，看实测符合哪一个：

```
k = H·tan(a)
V = H/6·(A_bot + 4·A_mid + A_top)     ← 棱台用 Prismatoid 公式是精确的，不是近似
ΔV = H·[±k(L+W) + 4k²/3]
```

40×30×20 拔 5°：上大下小 **+2531.328**，上小下大 **−2368.037**。

**两者大小并不相等**（差了 `2k(L+W)H`），所以"二选一"仍然是硬判据，不是放水。两个都对不上时 `matched_direction` 报 `null`——**不许在对不上的时候声称符合哪一个**。

### 15.5 冒烟测试：三段必须互相隔离

```powershell
python scripts\shape_features_smoke.py
```

三个特征各自新建一个长方体，各自 `try`。因为它们的失败概率是递增的，如果串在一个模型上，第一个失败就看不到后两个的证据了——**一轮真机往返只能修一个 bug**。分开跑，一次就能拿到三份独立结论。

另外，期望值在测试文件里**独立推一遍**，不 import 实现里的公式。测试和被测代码共用同一个错误公式，等于什么都没测。

Shell 那段还多了一条**独立于体积**的证据：核对被去掉面的重心 Z 是否等于 H、面积是否等于 L×W。万一某个别的面碰巧凑出相近体积，这条能立刻把"挑错面"和"壁厚算错"区分开。

### 15.6 首轮实测：预期错了一半，而且错得有价值

跑之前我押的是「Chamfer 大概率一次过，Draft 不指望」。实测：

```
❌ FAIL  chamfer
✅ PASS  shell
❌ FAIL  draft
```

Shell 一次过，而且过得很干净：

```
removed_face      : {'index': 0, 'area_mm2': 1200.0, 'cog_mm': (20.0, 15.0, 20.0)}
face_candidates   : 6
measured_removed  : 13871.999999999998   理论 13872.0   rel = 1.31e-16
```

**我以为最难的那关反而最顺**——因为它复用的是已经验证过的东西（重心读取、精确解验证），而不是在赌新的 API。

Chamfer 反倒栽了。三条报错说得很清楚：

```
AddNewChamfer(mode=1): (-2147352562, 'Invalid number of parameters.', ...)
AddNewChamfer(mode=2): 同上
AddNewChamfer(mode=0): 同上
```

注意 `AddNewSolidEdgeChamfer` **一条报错都没有**——因为它在这台机器上压根不存在，`getattr` 直接返回 `None` 被跳过了。而 `AddNewChamfer` 三次都是「参数个数不对」，说明真实签名比我写的多一个参数（应是 `CatChamferOrientation`）。

Draft 的失败更直白：`Part.CreateReferences` 不存在。

### 15.7 三个失败是同一个病因：拿二手文档当一手事实

```
AddNewSolidEdgeChamfer      这台机器上根本不存在
AddNewChamfer(5 个参数)     参数个数不对
Part.CreateReferences       getattr 直接是 None
```

我一直在"按名字试探"——猜一个名字，试试在不在；猜一个参数个数，试试报不报错。试探只能回答"**我猜的这个在不在**"。

但 **CATIA 的类型库就在那台机器里**。`ITypeInfo` 能把真实存在的方法名、参数个数、参数名一条条读出来，回答的是"**到底有哪些**"。这是根本不同量级的信息。

所以加了 `probe_shape_api()`：

```powershell
python scripts\probe_draft_api.py
```

它直接打印 ShapeFactory 上所有含 chamfer / draft / shell 的方法及其参数表，以及 `Part` 上所有 `Create*` 方法。和 `probe_export_formats` 一样是**诊断工具**，没有固定合格标准，所以不进回归。

> 教训写在这儿：**试探是问"是不是"，反射是问"是什么"。有反射可用的时候，别停在试探。**

### 15.8 Chamfer 改成自收敛：错了当场撤销，换下一个

光把参数个数改对还不够，因为还剩一个更阴的不确定量。

`AddNewChamfer` 的最后一个参数是"角度**或**第二条边长"，取决于 mode 枚举。而 mode 的取值我不确定。如果猜反了，把 `45` 当成 45mm 的边长——**调用完全合法**，只是做出来的东西不对。这正是之前 COG 返回 `(0,0,0)` 那类失败。

更麻烦的是：45° 斜角在两种解释下的**正确结果恰好是同一个形状**（两条腿都等于 d）。所以谁对谁错，读文档分辨不了，只能算体积。

于是把策略链升级成**自收敛**的：

```
对每个 (方法, 参数个数, mode, 朝向, 第二参数含义) 组合：
    建 → Update → 量体积
    对得上  → 采纳，把生效组合写进 strategy
    对不上  → 把这个特征从树上删掉，换下一个
```

关键是那个**删除**动作。没有它，策略链就只能试一次——第一个"被接受但做错"的组合会永久赖在树上，后面所有尝试都建在一个已经错了的模型上。

这样精确解的角色又变了：15.2 里说它是"枚举的判据"，现在它是**搜索的目标函数**。一轮真机往返就能收敛，不用来回贴报错。

### 15.9 Draft 准备了退化路径

`References` 集合造不出来，就**一个面建一个 Draft 特征**。

四个侧面用的是同一个中性面、同一个方向、同一个角度,分四次做和一次做出来的形状一样,只是特征树上多三个节点。

> **能接受的降级要说清代价。** 这里的代价就是树变长了,仅此而已——和 STEP 降级成 IGES 时说清"体积回读拿不到"是同一个规矩。半途失败还要把已建的清掉,否则下一轮尝试是在脏模型上做的。
