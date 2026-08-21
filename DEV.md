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

### 12.5 首次基线（CATIA V5R34 SP3，已实测）

```
  ✅ PASS       0.5s  connect   COM 链路 / 版本 / 会话
  ✅ PASS       4.0s  health    超时熔断 + 阻塞自诊断 + 重建恢复
  ✅ PASS       2.6s  box       加料特征（Sketch + Pad）
  ✅ PASS       1.9s  pocket    去料特征（偏移平面 + Pocket）
  ✅ PASS       2.7s  fillet    几何引用（边拾取 + EdgeFillet）
  ✅ PASS       5.2s  export    交付（STEP，不可用则降级 IGES）
  ✅ PASS       6.5s  mcp       MCP 端到端（起服务器 + 调工具）

  7/7 通过，合计 23.4s
```

**23 秒**——这个数字比「7/7 通过」更重要。回归的执行率由它决定：23 秒的回归每次改动都会跑，5 分钟的回归只在提交前跑，20 分钟的回归没人跑。所以以后加特征时，单个脚本的耗时要盯住。

基线里几个值得记的实测点：

- `health` 里那 4.0s 有 3.0s 是**故意等的**（人为制造卡死后验证熔断），不是性能问题。
- `export` 两次都落到 `format_used: igs`、`reimported_volume_mm3: null` —— 这台机器 STEP 翻译器没许可，降级路径按预期生效。**降级是已知结论，不是失败**，所以判定给 PASS。
- `fillet` 稳定在 `edges=12of16`、`relative_error=1.4e-09`，与单独跑时一致 —— 说明串跑没有污染它。

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

### 13.4 单位与出参数组：两个真实的坑

**单位**。CATIA 的 `Measurable` 走 SI —— 体积 m³、面积 m²、坐标 m。换算是 ×1e9 / ×1e6 / ×1e3。好在这个坑**不可能被漏掉**：错了会差整整 6 个数量级，冒烟测试一跑就炸。

**出参数组**。`GetCOG` 不是返回值，而是**出参数组**（`CATSafeArrayVariant`）。pywin32 传这种参数有两种写法，能不能用取决于它对 byref VARIANT 的支持：

```python
# 姿势 A（标准解法）
arr = win32com.client.VARIANT(
    pythoncom.VT_ARRAY | pythoncom.VT_BYREF | pythoncom.VT_R8, [0.0, 0.0, 0.0])
measurable.GetCOG(arr)
vals = arr.value

# 姿势 B（靠 pywin32 自动 marshalling）
buf = [0.0, 0.0, 0.0]
out = measurable.GetCOG(buf)
```

按序试，生效的那种记进 `cog_strategy`。换台机器出问题时，**这一个字段就能说明是不是 marshalling 姿势变了**，不用从头查起。

三项也是分开读、分开记错误的：体积测不出不该连累面积也没有——能拿到几条就报几条，剩下的说明为什么拿不到。

### 13.5 冒烟测试

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

**重心不变这一条是这个测试里最有意思的**：倒圆角对三个中心面都是对称的，所以重心必须原地不动。一旦漂移，说明有边没倒到或倒错了边——而这恰恰是**体积几乎测不出来**的那类错误（少倒一条边，体积只差百分之几，重心却立刻偏）。

### 13.6 让 AI 测量

对 AI 说：**"看看现在打开的这个零件，体积和重心是多少？"**

它应先 `inspect_document` 看清结构，再 `measure_body`。这是第一次 AI 能对**它没参与建的模型**给出有依据的回答。
