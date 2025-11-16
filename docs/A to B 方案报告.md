### 1. 核心交互：COM XML 操纵
*   **接口**: `Microsoft.Office.Interop.OneNote.IApplication`
*   **数据流**: `GetPageContent` (获取页面XML) -> 解析并定位选区节点 -> **本地转换逻辑** -> 构建新 XML 片段 -> `UpdatePageContent` (提交修改)。
*   **Schema 版本**: 强制使用 **`xs2013`** (`http://schemas.microsoft.com/office/onenote/2013/onenote`)，避免 `xsCurrent` 带来的结构不确定性。
*   **原子性**: `UpdatePageContent` 是基于 `ObjectID` 的整块替换。必须精确获取待替换段落的 `ObjectID`，否则会导致并发冲突或光标跳动。

### 2. Markdown 转换 (Markdig -> OneNote XML)
*   **解析库**: 使用 **Markdig** (C#) 解析 Markdown 为抽象语法树 (AST)。
*   **渲染策略**: 实现自定义 `IMarkdownExtension` 或 Visitor，不生成 HTML，而是直接构建 `XElement` (LINQ to XML) 对应的 OneNote XML 结构。
    *   **标题**: 映射为 `one:OE` 的 `quickStyleIndex` 属性（需预先在 XML 头部 `<one:QuickStyleDef>` 定义样式索引）。
    *   **列表**: 映射为嵌套的 `<one:Outline><one:OEChildren>...` 结构。
    *   **粗体/斜体**: 映射为 `<one:T><!]></one:T>`。**注意**：OneNote API 仅支持极少数内联 CSS (Color, Font-size, Bold, Italic)，不支持 class 或 div。

### 3. 代码块与语法高亮 (Code Block -> Table)
OneNote 无原生代码块对象，需模拟：
*   **容器模拟**: 将代码块转换为 **1x1 表格** (`one:Table`)。
*   **背景色**: 设置 `<one:Cell shadingColor="#F0F0F0">` 实现灰色背景。
*   **语法高亮**:
    1.  使用 **ColorCode-Universal** 或 **TextMateSharp** 离线解析代码文本。
    2.  将生成的 Token 颜色转换为 `<span style='color:#RRGGBB'>`。
    3.  所有内容包裹在 CDATA 中。

### 4. LaTeX 公式转换 (LaTeX -> OMML)
OneNote 仅支持 OMML (Office Math Markup Language)。转换路径如下：
1.  **LaTeX -> MathML**: 使用 **CSharpMath** 或 **AngouriMath** 库（纯 C# 实现）将 LaTeX 字符串解析为 Presentation MathML XML 字符串。
2.  **MathML -> OMML**: 利用 Office 安装目录自带的 **`MML2OMML.XSL`** 文件（通常在 `%ProgramFiles%\Microsoft Office\Office16\MML2OMML.XSL`）。
3.  **转换执行**: 使用.NET 的 `XslCompiledTransform` 类加载上述 XSLT，将 MathML 流转换为 OMML XML (`m:oMathPara`)。
4.  **注入**: 将生成的 OMML 节点直接插入 `one:OE` 节点中。

### 5. Mermaid 图表 (WebView2 -> Image)
由于 OneNote 无法渲染 JS，必须转为图片：
*   **渲染引擎**: 集成 **Microsoft Edge WebView2** 控件。
*   **离线方案**:
    1.  初始化一个 **不可见 (0x0尺寸)** 的 WebView2 实例。
    2.  注入本地 `mermaid.min.js` 和绘图 HTML。
    3.  JS 执行 `mermaid.render()`。
    4.  JS 计算渲染后的 DOM 尺寸 (`document.body.scrollHeight/Width`) 并回调 C# 调整 WebView2 视口大小（确保截图完整）。
    5.  调用 `CoreWebView2.CapturePreviewAsync` 截取 PNG 流。
*   **存储**: 将图片流转为 Base64，嵌入 `<one:Image><one:Data>BASE64...</one:Data></one:Image>` 节点。

### 6. 预览窗格与滚动同步
*   **预览实现**: 使用 **WebView2** 渲染 HTML。这是唯一的富文本预览方案。
*   **滚动同步 (Hack)**: OneNote API **不提供** `ScrollTo` 方法。
    *   *唯一解法*: 使用 `IApplication.NavigateTo(HierarchyID, ObjectID)`。
    *   *逻辑*: 在 Markdown 预览中点击某段落 -> 计算对应的 OneNote `ObjectID` -> 调用 `NavigateTo`。这会强制 OneNote 界面跳转并滚动到该对象位置。

### 技术栈汇总 (推荐)
*   **核心**:.NET Framework 4.8 或.NET 6+ (取决于 Office 版本兼容性要求)
*   **Markdown**: Markdig
*   **LaTeX**: CSharpMath + MML2OMML.XSL
*   **Mermaid**: WebView2 (Microsoft.Web.WebView2)
*   **XML 处理**: System.Xml.Linq (XDocument)
*   **UI**: WinForms / WPF (用于设置页和 WebView 容器)

功能模块,原始格式 (A),目标格式 (B),核心技术/库,关键实现点
文本排版,Markdown,OneNote XML,Markdig,AST 遍历 -> 构建 one:OE 结构
代码高亮,Code Block,1x1 Table + Spans,ColorCode-Universal,one:Cell shadingColor + CDATA Inline Styles
数学公式,LaTeX,OMML (XML),MML2OMML.XSL,LaTeX -> MathML (CSharpMath) -> OMML (XSLT)
流程图表,Mermaid,PNG Image,WebView2,Headless Browser -> CapturePreviewAsync -> Base64
API 交互,-,-,OneNote Interop,GetPageContent / UpdatePageContent (xs2013)