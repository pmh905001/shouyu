# 截图文字提取（OCR）集成方案

> 状态：**方案已确认，待实现**
> §6 的待讨论点已拍板：**自建 PySide6 选区遮罩**（§2.2）+ **直接定 RapidOCR**（§3，跳过 Windows OCR spike）+ **独立新热键**（不复用连按机制）。

## 1. 背景与目标

你截图保存重要信息时，很多时候真正想要的是**截图里的文字**，而不是一张图片本身（图片存进 Excel 后没法搜索、没法复制、没法二次编辑）。你的思路：

1. 选中屏幕上的一块区域。
2. 提取这块区域里的文字（OCR）。
3. 把提取到的文字放进系统剪贴板。
4. 触发现有的"保存剪贴板"热键——这一步完全复用你已经有的功能，不需要新写保存逻辑。

也就是说，这次要新增的只是"选区域 → 识别文字 → 写回剪贴板"这一段，后面全部借现成的（尤其是刚做完的 `message_queue`/`dispatch` 那套，OCR 出来的文字走的是和普通剪贴板文本完全一样的 `clipboard_append` 路径）。

## 2. 选区域怎么做

两种思路，都调研了：

### 2.1 借用 Windows 自带的 Snip & Sketch（最省事）

Windows 10/11 自带的截图工具（`Win+Shift+S`，或者用 `ms-screenclip:` 这个 URI 直接唤起）选完区域后，**会自动把截下来的图片放进系统剪贴板**——这正好是你现在 `ImageGrab.grabclipboard()` 已经会读的东西。

- 优点：零新增 UI 代码，微软已经把"选区域"这个交互做得很成熟（有网格线、放大镜、多显示器支持）。
- 缺点：流程是两步——按一下热键唤起 Snip & Sketch，选完区域后，还要**再按一次**（同一个或另一个）热键让 shouyu 去读剪贴板、跑 OCR、再写回剪贴板——不是"按一下热键，选完就自动完事"的单步体验。

### 2.2 自建一个 PySide6 全屏半透明遮罩，拖拽选区域

你的 `habit_dialog.py` 已经在用全屏 `QDialog`，这里可以照着同样的路子做一个"截图选区"遮罩窗口：按一下热键 → 弹出全屏半透明遮罩 → 鼠标拖出一个矩形 → 松手 → 用 `PIL.ImageGrab.grab(bbox=...)`（或 `QScreen.grabWindow`）截下这块区域 → 直接送去 OCR，不用真的经过系统剪贴板这一步。

- 优点：一个热键、一步到位，体验上跟"按热键→拖一下→完事"一致，不需要用户记两次操作。
- 缺点：要自己写选区交互（矩形拖拽、多显示器坐标换算、Esc 取消），工作量比 2.1 大，但都是 PySide6 里很常规的东西，风险不高。

> 💬 待讨论 A：两步（借 Snip & Sketch）还是一步（自建遮罩）？我倾向**自建遮罩**——你现在其他功能全是"一个热键搞定"的风格（截图保存、剪贴板保存都是一步），中途插一个"去用系统截图工具"的手动步骤会显得不一致，且后续维护也在你自己代码里，不依赖 `ms-screenclip:` 这个 URI 会不会被系统改名之类的风险。

## 3. OCR 引擎怎么选

调研了 5 个选项，重点看：**中文识别准确率**（你的截图大概率中英文混排，甚至是代码/终端输出）、**要不要额外装东西**、**打包成 PyInstaller 单文件 exe 会不会有坑**（`shouyu.spec` 已经在用 PyInstaller）、**是否需要联网**。

| 方案 | 中文准确率 | 依赖/体积 | PyInstaller 打包 | 联网 | 备注 |
|---|---|---|---|---|---|
| **Windows OCR**（`winsdk` 调用系统自带的 `Windows.Media.Ocr`） | 高（就是 Win11 "文本操作"/PowerToys 文字提取器背后那套引擎） | 零额外模型，`pip install winsdk` 一个包 | ⚠️ WinRT/COM 绑定，PyInstaller 打包历史上容易踩坑（需要正确带上 WinRT 元数据），需要先做个小 spike 验证 | 否，纯本地 | **`winsdk` 目前只有 beta 版**（最新 `1.0.0b10`，还没出正式 1.0），`pip install winsdk` 默认会因为"没有稳定版"直接报错，需要 `pip install --pre winsdk` 或在 requirements.txt 里钉死具体 beta 版本号 |
| **RapidOCR**（`rapidocr-onnxruntime`，基于 PaddleOCR 模型，跑在 ONNXRuntime 上） | 高（中文场景专门优化过） | 中等，纯 pip 装，模型文件几 MB～十几 MB | 友好，模型是普通文件，PyInstaller `datas` 里带上就行，社区案例多 | 否，纯本地 | 综合看是**最平衡**的选项 |
| **PaddleOCR** 原生 | 高 | 重，`paddlepaddle` 框架体积大 | 差，历史上 PyInstaller 打包坑最多，还经常首次运行联网下模型 | 首次可能联网下模型 | 不建议，太重 |
| **EasyOCR** | 中等（偏欧美语言优化，中文不是最强项） | 重，依赖 `torch` | 一般，torch 打包体积大 | 首次运行默认联网下模型（可离线预下好） | 不建议，体积/联网都不理想 |
| **Tesseract + pytesseract** | 中等偏弱（中文，尤其是复杂排版/代码截图） | 需要额外装系统级 `Tesseract-OCR` 二进制 + 语言包，不是纯 pip | 需要把 Tesseract 整个安装目录一起打包，多一层运维负担 | 否 | 最成熟、最多文档，但中文这块不是它的强项 |

**推荐：优先验证 Windows OCR（`winsdk`），备选 RapidOCR。**

- Windows OCR 是"免费"的——不占用你说的"轻量"这个诉求里的任何额外空间，识别效果也是本机能拿到的最好的（毕竟是系统自带、专门优化过的引擎），速度也快（原生调用，不是 Python 侧的模型推理）。风险点集中在**打包**这一件事上——建议先花很短时间写一个独立的小脚本（不进主程序）验证 `winsdk` 在这台机器上能正常调用、且能用 PyInstaller 打包出可运行的 exe，如果打包顺利，直接定这个方案；如果打包踩坑严重，退到 RapidOCR。
- RapidOCR 作为稳妥备选：万一 Windows OCR 打包不顺利，RapidOCR 不用担心 COM/WinRT 那一套坑，模型文件当成普通资源文件处理即可，跟你现有 `resources/icons/*.png` 走 PyInstaller 数据文件的方式是一样的套路，改动量可预期。

> 💬 待讨论 B：认可"先花小时间 spike 验证 Windows OCR 打包，不行再退到 RapidOCR"这个顺序吗？还是你想跳过 spike，直接定一个（比如你对折腾 WinRT 打包没耐心，就直接上 RapidOCR）？

## 4. 识别结果的排版问题（不管选哪个引擎都存在）

OCR 引擎返回的通常是"一堆文字块 + 每块的坐标"，不是天然排好序的一段文本。如果不处理，直接把识别出的文字块顺序拼起来，很容易把多行内容拼成乱序（尤其是代码/终端截图，缩进、对齐全乱）。需要按"先按行（y 坐标）分组，组内再按列（x 坐标）排序"的规则重新拼接，才能还原出接近原始排版的文本。这个逻辑跟引擎无关，选哪个方案都要写。

另外要如实设置预期：**OCR 对代码screenshot 的准确率天生比对自然语言文本低**——容易把 `0`/`O`、`1`/`l`、全角/半角符号搞混，这是所有 OCR 引擎的通病，不是选错引擎的问题。截的是代码/命令的话，建议提取完之后自己过一眼再用，不要完全信任。

## 5. 与现有系统的集成方式

复用你已经有的东西，新增的代码只是"前面这一小段"：

```
[热键] 触发选区域截图（§2 的方案）
    ↓
[OCR 识别] 得到文字（§3 的引擎 + §4 的排序逻辑）
    ↓
pyperclip.copy(识别出的文字)          # 按你的原话，先放回剪贴板
    ↓
Shortcut._enqueue_clipboard_save(column)   # 直接复用已有的入队逻辑
    ↓
（后面全部是已经做好的：message_queue 落盘 → toast「已记录」→ 后台 dispatch 写 Excel）
```

新增大约是：
1. 一个新热键配置项，比如 `kb.ini [shortcuts] ocr_capture=ctrl+alt+o`（沿用现有 `_add_hot_key_from_config` 的模式）。
2. 一个新的 `Shortcut.ocr_capture()` 方法：拉起选区遮罩 → 截图 → 调 OCR → 排序拼接 → `pyperclip.copy()` → 调用现有的 `_enqueue_clipboard_save`。
3. （如果选 §2.2 自建遮罩）一个新的 PySide6 全屏选区窗口，类似 `habit_dialog.py`/`todo_panel.py` 已有的窗口写法。

这条链路完全不需要碰 `message_queue.py`/`dispatch.py`——OCR 出来的文字就是一条普通的 `clipboard_append` 消息，队列/重试/落盘/通知这些都已经是现成的，不用因为这个新功能再改一遍。

## 6. 决策结论

- ✅ A：选区交互用**自建 PySide6 全屏半透明遮罩**（§2.2）。
- ✅ B：OCR 引擎**直接定 RapidOCR**（`rapidocr-onnxruntime`），跳过 Windows OCR 的打包 spike。
- ✅ C：绑定一个**独立新热键**（`kb.ini [shortcuts] ocr_capture=...`），不复用 `save_clipboard` 的连按机制。

## 7. 实现范围

1. `requirements.txt` 加 `rapidocr-onnxruntime`。
2. 新增 `shouyu/view/region_selector.py`：全屏半透明 `QDialog`，鼠标拖拽出矩形选区，Esc 取消，松手后返回选区的屏幕坐标（多显示器场景用 `QScreen` 的全局坐标）。
3. 新增 `shouyu/service/ocr.py`：`extract_text(image: PIL.Image) -> str`，内部用 RapidOCR 识别 + 按 §4 的"行优先、行内按 x 排序"规则拼接成文本。
4. `shouyu/action/shortcut.py` 新增 `Shortcut.ocr_capture()`：调起选区遮罩 → `PIL.ImageGrab.grab(bbox=选区)` 截图 → `ocr.extract_text()` → `pyperclip.copy(text)` → 调用现有的 `_enqueue_clipboard_save()`。
5. `kb.ini` 新增 `ocr_capture` 热键配置项 + 说明注释，`_add_hot_key_from_config` 里注册。
