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

### 5.4 再下一步

1. 加只读测量工具（`measure_*`），继续零风险扩充证据类工具。
2. 再加第一个**写操作**：新建 Part + 画长方体 Pad（进入"检查证据 → 迭代修复"闭环）。
3. 引入超时熔断 + 会话心跳看门狗（可行性分析里的 1、3 号风险）。
