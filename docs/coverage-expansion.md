# AnyView 文件类型覆盖扩展 — 设计稿

**作者**：Joe（owner）
**实现**：Weaver
**日期**：2026-04-18
**目标**：把 AnyView 能看的文件类型尽可能扩广，朝"任何文件丢进来都能看"靠拢。

## 现在的状态

四个渲染器：

| 渲染器 | 用什么实现 | 已覆盖格式 |
|---|---|---|
| PDFRenderer | PDFKit | pdf |
| ImageRenderer | NSImageView | png/jpg/gif/webp/tiff/bmp/ico/heic/svg |
| QuickLookRenderer | QLPreviewView | pptx/ppt/xlsx/xls/key/numbers/pages |
| WebRenderer | WKWebView + highlight.js + mermaid + docmod CLI | 60+ 代码语言、md、html、json/yaml/toml/xml/csv/plist/ini、docx/docmod/doct |

加新格式 = 给某个渲染器的 `supportedExtensions` 加扩展名，必要时给 Info.plist 加一段 UTI。

## 扩展原则

不重复造轮子。优先级按"所需工作量从小到大"分四层：

1. **Phase 1 — 白嫖 macOS 原生支持**：QuickLook 已经支持的格式，只需在 Info.plist 注册 UTI + 把扩展名加到 `QuickLookRenderer.supportedExtensions`
2. **Phase 2 — 在 WebRenderer 里加新文本/结构化格式**：用 WKWebView 渲染 HTML，写一个轻量转换器
3. **Phase 3 — 自定义渲染器，需要外部库或专门解析**
4. **Phase 4 — 大型专有格式，依赖重 / 商业库**：评估后再决定做不做

---

## Phase 1 — 白嫖（预计 1 天，含测试）

只改 Info.plist 和 `QuickLookRenderer.supportedExtensions`。每个格式实测能不能在 macOS QuickLook 里看到合理的预览，能就加上。

### 候选格式

| 类别 | 扩展名 | macOS QuickLook 支持情况 |
|---|---|---|
| 3D 模型 | stl, obj, usdz, usd, dae | macOS 有 SceneKit-based preview |
| 音频 | mp3, m4a, wav, flac, aac, aiff | QuickLook 显示波形 + 播放器 |
| 视频 | mp4, mov, m4v, mkv, avi, webm | QuickLook 内嵌播放器 |
| 字体 | ttf, otf, ttc, woff, woff2 | QuickLook 显示字符样张 |
| 压缩包列表 | zip, tar, gz | QuickLook 列出包内文件 |
| 邮件 | eml, msg | QuickLook 显示信头 + 正文 |
| 联系人 | vcf | QuickLook 显示联系人卡片 |
| 日历 | ics | QuickLook 显示事件 |

### 验收标准

- 每个加的格式：找一个真实样本文件，确认在 AnyView 里能打开、能看到内容（不是空白窗口或错误）
- 不能预览的从清单里移掉，不留半成品

---

## Phase 2 — WebRenderer 扩展（预计 1-2 周）

每个格式写一个"X 文件 → HTML 字符串"的转换函数，复用 WebRenderer 显示。

### 优先级排序（按用户痛感）

1. **Jupyter notebook (.ipynb)** — 用户多，结构清晰：JSON 里的 cells 转成 markdown / code block / output 的 HTML。代码段直接走现有 highlight.js，markdown 段走现有 markdown 渲染。预计 2-3 天。

2. **EPUB (.epub)** — 本质是 ZIP + 一组 XHTML。复用现有 ZipExtractor，把 OPF 里指定的第一个 spine 项展示。预计 3 天。

3. **GeoJSON (.geojson)** — JSON 里的 geometry 用 leaflet（CDN 引一份，或内嵌）展示在地图上。预计 2 天。

4. **MindMap (OPML, .mm)** — 折叠树。OPML 简单，FreeMind .mm 略复杂但格式开放。预计 3 天。

5. **字幕 (.srt, .vtt, .ass)** — 时间轴对齐文本。时间戳 + 文本两列。预计 1 天。

### 验收标准

- 每个格式：sample 文件能正确显示，复杂样本不崩溃
- 转换器单独可测（输入文件路径 → HTML 字符串）

---

## Phase 3 — 自定义渲染器（预计 2-4 周/格式）

需要新的 Renderer 实现 ViewerRenderer protocol，依赖外部库或自写解析。

### 候选

1. **DXF（CAD 2D 图纸）** — 用 [dxflib](https://github.com/codelibs/dxflib)（C++，通过 brew install dxflib 可装）解析，自己用 CoreGraphics 画到 NSView。或直接用纯 Swift 解析 DXF 文本格式（DXF ASCII 版本规范公开）。预计 1-2 周。
   - 备注：DWG 不在 Phase 3 范围（专有格式，需付费 ODA Teigha 库）

2. **Scientific data (.parquet, .hdf5, .npz)** — 不试图渲染数据本身，只显示 schema + 前 N 行。`.parquet` 可用 Apache Arrow C++ 库；`.hdf5` 可用 HDF5 C 库。预计 2 周/格式。

3. **DICOM (.dcm)** — 医学影像。需 DCMTK 或类似库。受众窄，**优先级低**，除非有具体用户需求。

4. **DjVu (.djvu)** — 扫描文档。需 DjVuLibre。受众窄，**优先级低**。

### 验收标准

- 每个：renderer 类自带 unit test（输入文件路径 → 渲染产物的某种快照）
- README 里加一行"现在 AnyView 也能看 X"

---

## Phase 4 — 不打算做的（除非有具体需求）

| 格式 | 原因 |
|---|---|
| DWG | 专有格式，开源解析不靠谱，商业库（ODA Teigha）需付费授权 |
| STEP / IGES（3D CAD） | OpenCascade 库 100MB+，依赖重，受众窄 |
| Adobe AI / PSD / INDD | 专有，没必要重新发明 |
| 早期 Word .doc（不是 .docx） | docmod CLI 已经覆盖不到，需 antiword 这类老库 |

每条如果你以后要加，单独提，重新评估。

---

## 推荐执行顺序

1. **本周内**：Weaver 跑 Phase 1，把 8 类约 30 个扩展名补上，AnyView 一夜之间能多看一倍格式
2. **下周开始**：Phase 2 第 1 项（ipynb），最有用户感知
3. **Phase 1 + Phase 2 第 1 项稳定后**：再决定 Phase 3 的第一个目标

每个 phase 结束做一次实测 demo（屏幕截图或者直接给 lizcc 看）。

---

## 谁做什么

- **Joe（我）**：本设计稿；Phase 1 完成后写一篇博客，介绍 AnyView 新覆盖范围；Phase 2 每个新格式上线前再加补充设计稿
- **Weaver**：实现 Phase 1（Info.plist + QuickLookRenderer 扩展）；评估 Phase 2 各项的具体实现路径，给出"能做 / 多久"表格
- **lizcc**：批准 Phase 1 启动；Phase 2 选第一项前 review 一下"用户需要哪个格式最强"

## 没想清楚 / 留给讨论

1. Phase 1 里 macOS QuickLook 的预览质量我没逐个验证过——可能有些格式实际打开是空白或卡顿。Weaver 实现时遇到 quality 问题可以缩小清单。
2. 视频/音频是否需要"全功能播放"还是只做静态预览？现在偏向只做 QuickLook 的内嵌播放，不自己写控件。
3. EPUB / Jupyter 这类多页文档是否需要导航条？现在 AnyView 没有"翻页/目录"概念，可能需要扩 ViewerRenderer protocol。
