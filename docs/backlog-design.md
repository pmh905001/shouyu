# Backlog（任务待办池）设计文档

> 状态：**讨论中（未开始实现）**
> 本文用于在动手前对齐存储结构、UI 与交互细节。文中带「💬 待讨论」的地方是需要你拍板/一起敲定的点。

## 1. 背景与目标

当前 `HabitDialog`（`shouyu/view/habit_dialog.py`）晨间界面是「今日习惯 | 今日要事」两列，并已有「昨日结转」卡片把昨天未完成的任务勾选带到今天。

痛点：没做完、又不属于"今天必须做"的任务，只能一直躺在某一天的 sheet 里或被遗忘。需要一个**跨天的任务池（Backlog）**：

- 未完成的任务可以放进 Backlog 暂存。
- 可以把 Backlog 里的任务**拖回当天**的任务列表。
- 任务区分**工作 / 生活**两类。

## 2. 已确认的需求决策

| # | 决策点 | 结论 |
|---|--------|------|
| 1 | 布局 | **三列**：`今日习惯 \| 今日要事 \| Backlog` |
| 2 | Backlog 分类展示 | **工作 / 生活 两段并列**（不是 Tab 过滤） |
| 3 | 收尾兜底 | 关闭时若今天有未完成任务，**弹提示询问是否移入 Backlog** |
| 4 | 拖拽范围 | **今日 ↔ Backlog 双向拖**，且 **Backlog 内部可拖动排序** |
| 5 | 存储 | Backlog 的工作/生活**分成不同的 Excel tab**（本次新增的想法，详见 §3） |

## 3. 底层存储设计（Excel）

### 3.1 现状回顾

- 一个工作簿，**每天一个 worksheet**，名字形如 `2026-07-25`（`excel.py:32`）。
- 另有特殊 sheet：`reflections`（`excel.py:482`）、`todo list`（`excel.py:117`）。
- 日 sheet 的「plan 区」列布局（`plan.py`）：
  - `B` 文本、`C` 时长、`D` 优先级、`E` 反思；状态用**字体颜色**编码（灰=待办 / 红=进行中 / 绿+删除线=完成）。
- 领域模型 `PlanTask`：`text / status / row / duration_minutes / priority / reflection`。

### 3.2 方案：工作/生活各用一个独立 sheet

按你的想法，Backlog 用**两个专用 sheet**（跨天全局，不按日期）：

- `backlog-work`（工作池）
- `backlog-life`（生活池）

**优点**
- 直接打开 Excel 看，工作/生活天然分开、清爽。
- 分类是**隐式的**（在哪个 sheet 就是哪类），backlog sheet 里不需要再存 category 列。
- 和 §4 的「两段并列」UI 一一对应，心智模型干净：一个 section ↔ 一个 sheet。

**代价 / 注意点**
- 工作↔生活「改分类」时，本质是把一行从一个 sheet 挪到另一个 sheet（删一行 + 加一行）。
- 今日 ↔ Backlog 移动时，要根据任务的 category 决定落到哪个 sheet；因此**日 sheet 仍需要一个 category 字段**（见 §3.4）。

> 💬 待讨论 A：sheet 命名。建议 `backlog-work` / `backlog-life`（英文、带连字符，和 `todo list` 风格接近）。也可用中文 `待办池-工作`。你倾向哪个？

### 3.3 Backlog sheet 列布局

`backlog-work` 与 `backlog-life` **列结构完全一致**：

| 列 | 含义 | 说明 |
|----|------|------|
| A | 文本 | 任务内容；字体颜色沿用 pending 灰色即可（池中默认都是待办） |
| B | 时长（分钟） | 空=未估时 |
| C | 优先级 | `P1/P2/P3`，空=无 |
| D | 备注/反思 | 复用 `PlanTask.reflection` |
| E | 入池日期 | `YYYY-MM-DD`，用于展示「已搁置 N 天」 |

- **第 1 行**可留作表头（`A1="backlog"` 之类，便于识别），任务从第 2 行开始，参考 `plan.py` 的 `PLAN_FIRST_TASK_ROW=2`。
- **行顺序 = 显示顺序**。Backlog 内部拖拽排序 → 保存时按新顺序**全量重写**该 sheet（复用 `write_plan_tasks` 的"重写 + 清理多余行"套路，`plan.py:206`）。

> 💬 待讨论 B：是否需要「入池日期」这列（E）？它能支撑"这条已经搁置 12 天了"的提醒，帮助清理僵尸任务。我建议**保留**，成本很低。

### 3.4 日 sheet 增加 category 列

日 sheet 的 plan 区新增一列 **`F` = category**（`work`/`life`）：

- 今日列表里每条任务带分类，才能显示工作/生活徽标、并在移入 Backlog 时路由到正确的 sheet。
- 新增 `PLAN_CATEGORY_COLUMN = "F"`，在 `read_plan_tasks` / `write_plan_tasks`（`plan.py:172` / `206`）读写。
- 旧数据该列为空 → 默认 `work`（见 §3.7 迁移）。

### 3.5 领域模型

新增分类枚举（放 `plan.py`，与 `TaskPriority` 并列）：

```python
class TaskCategory(str, Enum):
    WORK = "work"
    LIFE = "life"

    @classmethod
    def from_value(cls, raw) -> "TaskCategory":
        # 空/未知 -> WORK
        ...
```

模型复用策略（💬 待讨论 C）：

- **方案 C1（推荐）**：给 `PlanTask` 增加两个字段 `category: TaskCategory = WORK` 和 `created_date: str = ""`。今日任务与 Backlog 任务共用一个 dataclass，拖拽来回搬时**无需类型转换**，只是切换所在的列表容器。`created_date` 仅 Backlog 用到（今日任务留空）。
- 方案 C2：单独 `BacklogTask` dataclass，语义更纯，但每次今日↔Backlog 拖拽都要转换字段，代码更啰嗦。

我倾向 **C1**：字段少、转换零成本，`created_date` 对今日任务只是恒空而已。

> 若采用 C1，需要**同步更新所有 clone/snapshot 点**，否则拖拽/撤销会丢字段：
> `_read_yesterday_snapshot`（`habit_dialog.py:305`）、`_clone_tasks`（`:1048`）、`_snapshot_tasks`（`:1604`）、`_carry_over_now` 里的 `PlanTask(...)`（`:1353`）。

### 3.6 服务层与持久化

- 新增 `BacklogService`（放 `plan.py` 或新建 `backlog.py`）：`read() / write(tasks) / add(task) / remove(task)`，绑定某个 backlog sheet，实现参考 `PlanService`。
- `KbExcel` 增加 `backlog_service(category) -> BacklogService`，**惰性创建** sheet（参考 `stage_reflection` 的 `create_sheet` 写法，`excel.py:482`）。
- **保存时机**：Backlog 是全局数据，Dialog 内持有工作副本。保存时把两个 backlog sheet **按内存顺序全量重写**。接入现有后台保存流程 `_persist_plan_in_background`（`habit_dialog.py:157`）的 `_stage_changes` 里一起 stage，复用它的**失败重试 + `.unsaved` 兜底**机制，无需另写保存逻辑。

> 💬 待讨论 D：Backlog 全局副本的并发/新鲜度。Dialog 打开时读一次两张 backlog sheet 到内存；期间若你在 Excel 里手改了 backlog，保存时会被内存副本覆盖。现状 plan 区也是这个模型（打开即快照），保持一致即可。可接受吗？

### 3.7 迁移与兼容

- 全部**惰性创建、零迁移脚本**：旧工作簿没有 `backlog-*` sheet、日 sheet 没有 `F` 列 → 首次用到时创建/按默认值读取。
- 日 sheet `F` 列为空 → category 读作 `WORK`。
- Backlog sheet 不存在 → 读出空列表。

## 4. UI 设计（三列）

### 4.1 整体布局

`_build_ui`（`habit_dialog.py:454`）现在是 `QSplitter(习惯卡 | 今日要事卡)`。改为三段：

```
QSplitter(Horizontal)
├── 习惯卡        stretch 1
├── 今日要事卡     stretch 2
└── Backlog 卡     stretch 2   ← 新增
```

全屏对话框宽度足够；stretch 比例可调。

### 4.2 Backlog 卡内部（工作/生活并列）

卡内用一个**纵向 QSplitter** 上下两段，可拖拽调整高度：

```
Backlog 卡
├── 🏢 工作 (N)        [+ 新增]
│   └── QListWidget  (backlog-work)
└── 🏠 生活 (M)        [+ 新增]
    └── QListWidget  (backlog-life)
```

- 每段一个 header：图标 + 分类名 + 数量徽标 + 「+ 新增」按钮。
- 每段一个 `QListWidget`，复用今日列表的渲染 `_format`（`habit_dialog.py:94`），但：
  - 池中不显示状态字形（都是待办），或统一显示 `○`。
  - 追加「⏳ Nd」入池天数（由 `created_date` 算，仅当 > 1 天时显示，避免噪音）。
  - 保留优先级 🔴/🟡/⚪ 和时长 ⏱ 徽标。

### 4.3 今日列表的分类呈现

- 今日列表每条任务前/后加分类徽标：`🏢` / `🏠`（在 `_format` 里加）。
- `_update_stats`（`habit_dialog.py:1412`）统计里加一句「工作 X · 生活 Y」。
- 今日列表**不做分组**，只用徽标区分（分组会显著增加渲染复杂度，先不做）。

## 5. 交互细节

### 5.1 拖拽矩阵

参与拖拽的共 3 个 `QListWidget`：`今日`、`工作池`、`生活池`。

| 从 → 到 | 行为 |
|---------|------|
| 今日 → 工作池 / 生活池 | 从今日移除，加入目标池；category 设为目标池的分类 |
| 工作池 / 生活池 → 今日 | 从池移除，加入今日 plan（状态=待办，带上原分类） |
| 工作池 ↔ 生活池 | **改分类**：从源池删除、加入目标池（跨 sheet 移动） |
| 池内 / 今日内 拖动 | 纯排序（现有今日列表已支持 InternalMove，`habit_dialog.py:628`） |

### 5.2 拖拽实现（关键难点）

`QListWidget` 用 `item.setData(_TASK_ROLE, task)` 存的 `PlanTask` 对象**默认不会随拖拽 MIME 序列化**，跨 widget 拖过去会丢。采用 **共享拖拽载荷**方案（此前已定）：

1. 三个列表都设 `setDragDropMode(DragDrop)` + `setDefaultDropAction(Qt.MoveAction)`。
2. `startDrag` 时把被拖 `PlanTask` 引用存到 `self._drag_payload` 和来源标识 `self._drag_source`。
3. 重写目标列表 `dropEvent`：读 `self._drag_payload`，插入目标内存列表、从来源内存列表删除、按目标分类改 `category`，最后 `_render_*` 重绘。
4. drop 完成后 `_push_undo()`（参考 `_on_rows_moved`，`habit_dialog.py:1171`），纳入撤销。

> 💬 待讨论 E：内部排序沿用 Qt `InternalMove`、跨 widget 走自定义 `dropEvent`，需要在同一组件上兼容两种模式。实现上把三个列表统一成"自定义 drop"处理最稳（内部移动也自己算 index），避免 InternalMove 与跨 widget 冲突。同意这个取舍吗？

### 5.3 右键菜单等价入口（不依赖拖拽也能操作）

在 `_show_context_menu`（`habit_dialog.py:1221`）扩展：

- 今日任务右键：「📥 移入 Backlog → 🏢 工作 / 🏠 生活」。
- Backlog 任务右键：「⬆ 拉到今日」「🔀 改为 生活/工作」「编辑 / 删除 / 优先级 / 时长」。

### 5.4 未完成任务清理（手动按钮，非自动弹窗）

> **最终决策（已修订）**：早期实现是"关闭时自动弹窗询问是否移入 Backlog"，但由于该界面**经常被打开只为查看当日任务**，每次关闭都弹窗非常打扰。故改为**纯手动按钮**，不再在关闭时自动提示。

- 在「今日要事」标题行放一个按钮 **「清理未完成 → Backlog」**（`_build_task_card`）。
- 点击 `_sweep_unfinished_to_backlog`：把今天所有**未完成**（pending / in_progress）任务按各自 category 移入对应 backlog pool，按文本去重；已完成（done）任务**保留在今天**作为当天记录。
- 该操作是显式的用户动作，且**可 Ctrl+Z 撤销**，所以不需要确认弹窗。今天没有未完成任务时给一个轻量成功提示。
- `reject()`（Esc / ✕关闭 / 跳过）**不再**触发任何 Backlog 相关弹窗，只保留原有的"未保存变更确认"。

### 5.5 与「昨日结转」的关系

现有「昨日结转」卡片（`habit_dialog.py:729` `_render_yesterday`）从**昨天的 sheet**捞未完成任务。引入 Backlog 后有两条演进路线：

- **路线 1（本期，低风险）**：Backlog 与「昨日结转」**并存互不干扰**。昨日结转照旧；Backlog 是独立的长期池。
- 路线 2（后续）：把「昨日未完成」也自动沉淀进 Backlog，晨间只保留 Backlog 一个入口，统一心智。

> 💬 待讨论 G：本期先走**路线 1**（不动昨日结转逻辑），后续再考虑合并。同意吗？

### 5.6 撤销 / 重做

所有 Backlog 变更（拖入拖出、改分类、池内排序、新增/删除）都要 `_push_undo()`，但**注意**：现有 undo 栈只快照 `self._tasks`（今日）。引入 Backlog 后需要把**三个列表一起纳入快照**，否则撤销一次跨列表拖拽会状态错乱。

> 💬 待讨论 H：撤销范围要扩到"今日 + 工作池 + 生活池"三者一起快照。这会改动 `_clone_tasks/_undo/_redo/_push_undo`（`habit_dialog.py:1047-1108`）。确认要做（否则拖拽撤销不安全）？

## 6. 边界情况

- **去重**：拖入 Backlog 或拉回今日时，同分类内按文本去重（参考 `_carry_over_now` `:1342` 的去重逻辑）。
- **空文本行**：读写时跳过（沿用 `_snapshot_tasks` `:1604` 的过滤）。
- **入池天数**：`today - created_date`；跨池移动是否重置 `created_date`？建议**不重置**（保留原始搁置时长更有意义）。💬 待讨论 I。
- **done 任务**：不进 Backlog；只有未完成任务可入池。
- **反思字段**：随任务一起带走（保留），池中可编辑。

## 7. 待讨论问题汇总

| 编号 | 问题 | 我的建议 |
|------|------|----------|
| A | backlog sheet 命名 | `backlog-work` / `backlog-life` |
| B | 是否保留「入池日期」列 | 保留（成本低，能治僵尸任务） |
| C | 模型复用 vs 独立 BacklogTask | 复用 `PlanTask` + 加 `category`/`created_date` 字段 |
| D | 全局副本新鲜度（打开即快照） | 与 plan 区一致，可接受 |
| E | 拖拽统一走自定义 dropEvent | 同意为宜（最稳） |
| F | 未完成任务清理方式 | **改为手动按钮**「清理未完成 → Backlog」，关闭时不再自动弹窗（详见 §5.4） |
| G | 与昨日结转的关系 | 本期并存（路线 1），后续再合并 |
| H | 撤销扩到三列表一起快照 | 需要做，否则拖拽撤销不安全 |
| I | 跨池移动是否重置入池日期 | 不重置 |

## 8. 预估改动文件清单（实现阶段用）

- `plan.py`：`TaskCategory`、`PlanTask` 加字段、plan 区 `F` 列读写、`BacklogService`（或拆到 `backlog.py`）。
- `excel.py`：`backlog_service(category)` + 惰性建两张 sheet；`_persist_plan_in_background` 里 stage backlog。
- `habit_dialog.py`：三列布局、Backlog 卡（工作/生活并列）、三列表拖拽、右键项、category 徽标、收尾兜底提示、undo 扩到三列表、所有 clone/snapshot 点补字段。
- `config.py` / `kb.ini`：可选开关（如默认分类、是否开启兜底提示、入池天数提醒阈值）。

## 9. 建议的分阶段实施（讨论定稿后）

1. **存储层**：`TaskCategory` + `PlanTask` 字段 + plan 区 F 列 + `BacklogService` + 两张 sheet 惰性创建 + 保存接入。
2. **只读 UI**：三列布局 + Backlog 卡渲染（先能显示两张池的数据）。
3. **非拖拽操作**：右键「移入/拉回/改分类」，先跑通数据流。
4. **拖拽**：三列表自定义 drop + 内部排序 + undo 扩展。
5. **收尾兜底**：关闭时提示移入 Backlog。
6. 打磨：入池天数提醒、统计、去重、边界。

---

**下一步**：请就 §7 表格里 A~I 逐条确认或修改；定稿后我再按 §9 分阶段实现。
