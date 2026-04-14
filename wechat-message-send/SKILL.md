---
name: wechat-message-send
description: Use when the user wants to find personal WeChat wxid/chatroom ids, resolve a chat by keyword or group name, remember a stable id-to-keyword mapping, or send a WeChat message through the Windows desktop client.
metadata:
  openclaw:
    emoji: "💬"
    os:
      - "win32"
    requires:
      anyBins:
        - "python"
        - "py"
---

# WeChat Message Send Skill

This skill wraps a local Windows WeChat workflow into one importable package.

It supports the full chain:

1. Scan local WeChat evidence to collect `wxid_*` and `*@chatroom`
2. Search IDs by keyword or nearby context
3. Remember stable `target_id -> search keyword` mappings
4. Send messages through the desktop WeChat client using stable UI automation

## Scope

This skill is for:

- personal desktop WeChat on Windows
- the installed and running `WeChat.exe`
- sending messages through UI automation, not injection

This skill is **not** for:

- direct protocol sending by internal id alone
- cloud APIs
- enterprise WeCom sending

## Files

- Primary workflow entrypoint: `scripts/wechat_skill_runner.py`
- Low-level id scan tool: `scripts/wechat_id_tool.py`
- Low-level sender: `scripts/wechat_sender.py`
- Persistent mappings: `data/target_mappings.json`

## Core Rule

Prefer the most stable route:

1. If the user says "给某某某发消息" and provides a name / remark / group name:
   use `send-by-name`
2. If the user gives an internal id and this skill already knows its search keyword:
   use `send-by-id`
3. If the user gives an internal id but no mapping exists:
   first `remember` a keyword for that id, then `send-by-id`
4. If the user says "给当前打开的聊天发":
   use `send-current`
5. If the user wants to find or verify an id first:
   use `scan` then `find`

## Message Encoding

Do **not** pass Chinese text directly on the shell command line when avoidable.

Always prefer writing the message to a UTF-8 temp file first, then passing `--message-file`.

Reliable pattern:

```powershell
@'
from pathlib import Path
Path(r'{baseDir}\data\message.txt').write_text('你好', encoding='utf-8')
'@ | python -
```

Then send with:

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" send-by-name --keyword "<联系人备注或群名>" --message-file "{baseDir}\data\message.txt" --wechat-path "<你的WeChat.exe路径>"
```

## Commands

Run commands from the skill root unless the host app sets another cwd.

### 1. Scan IDs

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" scan --wechat-path "<你的WeChat.exe路径>"
```

### 2. Find an ID by keyword

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" find "<群名关键词>" --kind chatroom
python "{baseDir}\scripts\wechat_skill_runner.py" find "张三"
```

If multiple strong candidates appear, do not guess silently. Show the top candidates and ask one short clarification question.

### 3. Remember a stable mapping

Use this after you have confirmed which conversation corresponds to a target id.

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" remember --id "<chatroom-id或wxid>" --keyword "<联系人备注或群名>"
```

### 4. Send by name / remark / group name

Preferred for natural requests such as "给某某某发消息".

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" send-by-name --keyword "<联系人备注或群名>" --message-file "{baseDir}\data\message.txt" --wechat-path "<你的WeChat.exe路径>"
```

### 5. Send by remembered ID

Preferred when the user refers to a known `wxid` or `chatroom id`.

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" send-by-id --ids "<chatroom-id或wxid>" --message-file "{baseDir}\data\message.txt" --wechat-path "<你的WeChat.exe路径>"
```

### 6. Send to the currently open chat

Use only when the user explicitly means the already-focused conversation.

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" send-current --message-file "{baseDir}\data\message.txt" --wechat-path "<你的WeChat.exe路径>"
```

## Safe Operating Procedure

Before sending:

1. Make sure WeChat is already logged in and visible
2. Make sure the user did ask to send the message
3. If using `send-by-name`, ensure the keyword is specific enough
4. If using `send-by-id`, ensure the mapping is already remembered
5. If the user wants a preview only, add `--dry-run`
6. If the command returns `ambiguous`, ask the user to choose a `pick_index`

During sending:

- The script will focus the WeChat window
- The script will control the keyboard and mouse focus
- Do not rely on other apps staying focused during the send
- The script will refuse to send if search results are ambiguous
- The script will refuse to send if the opened chat title does not match the expected target

After sending:

- Report that the send action was executed
- If you did not independently verify the visible UI result, say so briefly

## Decision Guide

When the user says:

- "给某某某发消息"
  - use `send-by-name`
- "给这个群 chatroom id 发消息"
  - use `send-by-id`
- "先查这个群名对应的 ID"
  - use `find`
- "把这个群记住以后就按 id 发"
  - use `remember`
- "给当前打开的微信聊天发"
  - use `send-current`

## Mapping Template

This package ships with an empty mapping file so it can be safely shared.

To teach the skill a stable id mapping on a new machine:

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" remember --id "<chatroom-id或wxid>" --keyword "<联系人备注或群名>"
```

## Notes

- A user can indeed be targeted by name / remark / group name with this skill, as long as WeChat search can uniquely locate the conversation
- Internal id discovery and stable sending are both supported, but stable sending by id depends on a remembered keyword mapping because this skill uses UI automation instead of injection

---

## Qclaw / 中文操作说明（Windows）

### 路径占位说明

- 技能目录示例：`C:\Path\To\wechat-message-send\`
- 微信客户端示例：`C:\Path\To\WeChat.exe`

将下面命令里的 `{baseDir}` 换成技能根目录（上两行之一即可）。

### 行为概要（与你的产品目标一致）

1. **唤起微信**：查找 `WeChat.exe` 进程与主窗口；若窗口最小化，会 `ShowWindow(SW_RESTORE)` 再 `SetForegroundWindow`，尽量抢到前台。
2. **搜索会话**：优先通过 UIA 查找微信 `Edit(title="搜索")` 搜索框并点击；若失败再退回 OCR / 比例坐标。输入关键词后会校验搜索框内容是否与关键词一致，避免误发。
3. **解析结果**：对搜索结果区域做 OCR，聚合成候选行；关键词唯一或完全匹配时自动选中；**多条匹配**时返回 `status=ambiguous`，JSON 里每条候选带 `pick_index`（1..N），**不发送**。
4. **过滤非联系人结果**：若 OCR 识别到的是“搜索网络结果”等面板，而不是“联系人 / 群聊”区块，则返回 `not_found`，**不发送**。
5. **打开聊天后校验**：点击候选进入会话，再 OCR 聊天区**顶部标题栏**；与期望昵称比对（含模糊匹配）；不匹配则 `title_mismatch` 并**不发送**。
6. **发送**：焦点到输入框，剪贴板粘贴正文，回车发送。

说明：搜索结果与标题 OCR 使用**屏幕坐标 BitBlt**截取（与肉眼所见一致，减轻 DPI / DWM 与 `PrintWindow` 不一致导致的「认对字却点错行」）；进程启动时调用 `SetProcessDPIAware`。结果列表按 OCR 框的**底边到下一框顶边**间距聚合成行。若打开会话后标题与候选昵称仍不匹配，会在同一流程内用**若干垂直像素偏移**重复搜索—点击—校验，仍失败则 `title_mismatch`。底层为 **RapidOCR**；搜索框优先依赖 **pywinauto / UIA**，缺失时会自动尝试安装。

### 退出码（供 Qclaw / 脚本判断）

| 退出码 | 含义 |
|--------|------|
| 0 | 成功执行发送，或计划内无阻塞（见 stdout 末行 `skill_exit=ok`） |
| 2 | 异常 / `SendError`（如未找到微信进程、无法置前窗口） |
| 3 | 需用户介入：stdout 末行 `skill_exit=needs_user_action`，且 JSON 行中 `status` 为 `ambiguous`、`not_found`、`title_mismatch`、`pick_out_of_range` 之一 |

### 多条搜索结果时：让用户选第几个

第一次执行若返回 `ambiguous`，stdout 中会打印候选列表（`name`、`contexts`、`pick_index`）。请用户选定后**用同一关键词**再执行，并加上 `--pick-index N`：

```powershell
python "{baseDir}\scripts\wechat_skill_runner.py" send-by-name --keyword "<联系人备注或群名>" --pick-index 2 --message-file "{baseDir}\data\message.txt" --wechat-path "<你的WeChat.exe路径>"
```

`send-by-id` 在按映射搜索时同样支持 `--pick-index`（对每个 `search` 模式目标生效）。

### 给 Qclaw 的提示词模板

**单目标发送（关键词尽量唯一）：**

```text
使用 wechat-message-send：用 --wechat-path "<你的WeChat.exe路径>"，对关键词「联系人备注或群名」发送消息（正文写入 data\message.txt 后用 --message-file）。若命令退出码为 3，读取 JSON 里 ambiguous/not_found/title_mismatch 原因回复我，不要猜测发送对象。
```

**预期可能多条时：**

```text
先按关键词「联系人备注或群名」准备发送；若返回 ambiguous，把候选列表和 pick_index 列给用户，等用户回复序号后再用相同关键词加 --pick-index 发送。
```
