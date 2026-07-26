---
name: record-changes-to-readme
description: Record every code change made in the shouyu repo into the "更新日志 (Changelog)" section of README.md. Use automatically after completing any code change (feature, fix, refactor, config, behavior tweak) to this repository, before ending the turn.
---

# Record Changes to README Changelog

After finishing any code change in this repo, append a changelog entry to `README.md` so the user does not have to ask each time.

## When to apply

Apply after you have made and verified a code change (added/edited/deleted code, config, or behavior) in this repo — do it as the final step of the turn, before your summary reply. Skip only for pure questions, read-only exploration, or changes to `README.md` itself.

## Where to write

The changelog lives under the `# 更新日志 (Changelog)` heading in `README.md`.

- Entries are grouped by date, **newest date on top**.
- Within a date, **newest entry on top**.
- Date format: `## YYYY-MM-DD` (use today's date).
- If today's date heading already exists, add a bullet under it; otherwise create a new date heading directly under the `# 更新日志 (Changelog)` intro line.

## Entry format

One bullet per logical change (not per file). Chinese, concise:

```
- **<简短标题>**：<做了什么 + 为什么/效果>。（<涉及文件，逗号分隔>）
```

Example:

```
- **番茄钟「去休息」提前结束**：专注阶段新增按钮，任务提前做完可直接进入休息，计一个 🍅 并按实际时长记录。（`shouyu/service/pomodoro.py`、`shouyu/view/pomodoro_window.py`）
```

## Rules

- File paths use forward slashes and are wrapped in backticks.
- Do not touch the corrupted mojibake intro section at the top of `README.md`.
- Group related file edits into a single bullet describing the user-facing change.
- Keep it factual — describe what changed and its effect, not implementation minutiae.
