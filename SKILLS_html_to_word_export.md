---
name: html-to-word-export
description: |
  当用户需要在纯前端（无后端）HTML 页面中实现以下任意功能时使用本 skill：
  1. 将富文本内容（含 LaTeX 数学公式、图片、多级标题）导出为可下载的 Word (.docx) 文件；
  2. 将同一页面内容导出为 PDF；
  3. 构建带侧边栏样式面板的单页文档编辑器 UI（字体/字号/行距/页眉/页码实时预览）；
  4. 在 docx.js 中处理 OMML 公式、tab stop 页眉、分页等高级排版需求；
  5. 使用 KaTeX 把 LaTeX 转为 MathML 再递归转换为 OMML（Word 原生公式格式）。
  涵盖 docx.js 8.x、FileSaver.js、KaTeX 0.16.x、html2pdf.js 的完整集成与已知 Bug 修复。
---

# SKILL: 从 HTML 网页生成可直接下载的 Word (.docx) 文档

> 适用场景：单页 HTML 应用、学术报告、科技文档，需在浏览器端（无后端）将富文本 + LaTeX 公式导出为规范 Word 文件。

---

## 一、总体架构与依赖

### 1.1 整体数据流

```
HTML 网页（编辑/预览）
    │
    ├─ 用户点击「下载 DOCX」
    │
    ▼
exportToDocx()  ← 纯前端，运行在浏览器
    │
    ├─ 1. 等待 docx.js + FileSaver.js CDN 加载完成
    ├─ 2. 读取当前样式设置（字体/字号/行距）
    ├─ 3. 遍历 CONTENT 数组，将每个内容块转成 docx 段落对象
    │       ├─ 普通段落  → Paragraph + TextRun
    │       ├─ 标题      → Paragraph + TextRun (bold, spacing)
    │       ├─ 含公式段落→ parseMathSegments → TextRun + Math (OMML)
    │       ├─ 独立公式  → buildMathForLatex → Math (OMML)
    │       └─ 图片      → buildImagePara → ImageRun + 图注
    ├─ 4. 构建 Document（页面尺寸、页边距、页眉、页码）
    ├─ 5. Packer.toBlob(doc)
    └─ 6. FileSaver.saveAs(blob, 'xxx.docx')
```

### 1.2 关键依赖库（CDN 方式引入）

| 库 | 版本 | 作用 | CDN |
|---|---|---|---|
| **docx.js** | 8.x | 构建 `.docx` 的所有元素（段落、文字、公式、页眉等） | `https://unpkg.com/docx@8/build/index.js` |
| **FileSaver.js** | 2.x | `saveAs(blob, filename)` 触发浏览器下载 | `https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js` |
| **KaTeX** | 0.16.x | LaTeX → MathML 解析（用于后续转 OMML） | `https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js` + auto-render |
| **html2pdf.js** | 0.10.x | HTML → PDF 截图导出 | `https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.1/dist/html2pdf.bundle.min.js` |

> **加载顺序陷阱**：docx.js 和 FileSaver 体积大，常在 `window.onload` 触发时还未完成加载。
> 必须用轮询等待：
> ```js
> const waitLib = (check, label, ms=15000) => new Promise((res, rej) => {
>   if (check()) { res(); return; }
>   const start = Date.now();
>   const t = setInterval(() => {
>     if (check()) { clearInterval(t); res(); }
>     else if (Date.now()-start > ms) { clearInterval(t); rej(new Error(`${label} 加载超时`)); }
>   }, 200);
> });
> await waitLib(() => !!(window.docx && window.docx.Document), 'docx.js');
> await waitLib(() => typeof saveAs === 'function', 'FileSaver');
> ```

---

## 二、内容层：CONTENT 数组与块类型

`CONTENT` 是页面内容的中心数据模型，每个元素描述一个逻辑块，同时驱动 HTML 预览渲染和 DOCX/PDF 导出。

### 2.1 块类型总览

| `type` 值 | 含义 | 关键字段 |
|---|---|---|
| `'h1'` / `'h2'` / `'h3'` | 标题（一/二/三级） | `text` |
| `'p'` | 正文段落（可含 `$...$` 公式） | `text`, `indent` |
| `'math'` | 独立显示公式（居中，不含其他文字） | `text`（纯 LaTeX） |
| `'img'` | 图片块 | `src`, `widthPt`, `heightPt`, `caption` |

### 2.2 文本块与样式同步（CSS 变量 ↔ docx 属性）

HTML 预览使用 CSS 变量实时更新，导出时读取同一套值：

```js
// HTML 侧：CSS 变量
document.getElementById('document-page').style.setProperty('--body-size', '12pt');

// 导出侧：从 <select> 读取相同值
const bodySize = ptToHalfPt(document.getElementById('s-bsize').value);  // "12pt" → 24
const bodyFontName = getDocxFontName(document.getElementById('s-bfont').value);
// "SimSun,'宋体',serif" → "SimSun"
```

字体名称对应关系（中文 Word 字体）：

| CSS 名称 | docx 字体名 |
|---|---|
| `'Microsoft YaHei'` | `Microsoft YaHei` |
| `'SimSun'` / `'宋体'` | `SimSun` |
| `'SimHei'` / `'黑体'` | `SimHei` |
| `'KaiTi'` / `'楷体'` | `KaiTi` |
| `'FangSong'` / `'仿宋'` | `FangSong` |

### 2.3 公式块

#### 2.3.1 内联公式与显示公式的分段解析

段落文字中混合文本和 `$...$` / `$$...$$`，需先分段再分别处理：

```js
function parseMathSegments(text) {
  const segs = [];
  let i = 0;
  while (i < text.length) {
    // 先检测 $$ (双美元，display math)
    if (text[i] === '$' && text[i+1] === '$') {
      const end = text.indexOf('$$', i+2);
      if (end !== -1) { segs.push({math:true, s:text.slice(i+2,end)}); i=end+2; continue; }
    }
    // 再检测 $ (inline math)
    if (text[i] === '$') {
      const end = text.indexOf('$', i+1);
      if (end !== -1) { segs.push({math:true, s:text.slice(i+1,end)}); i=end+1; continue; }
    }
    // 普通文本
    const next$ = text.indexOf('$', i);
    const chunk = next$ === -1 ? text.slice(i) : text.slice(i, next$);
    if (chunk) segs.push({math:false, s:chunk});
    if (next$ === -1) break;
    i = next$;
  }
  return segs;
}
```

然后：
```js
for (const seg of segs) {
  if (seg.math) {
    children.push(new D.Math({ children: latexToOmml(seg.s) }));
  } else {
    children.push(new D.TextRun({ text: seg.s, ... }));
  }
}
```

#### 2.3.2 LaTeX → OMML 完整转换流程

这是整个功能最复杂的部分，分三层：

```
LaTeX 字符串
    │
    ▼  (KaTeX)
MathML DOM（标准 XML）
    │
    ▼  (mathmlNodeToOmml 递归遍历)
docx OMML 元素数组
    │
    ▼
new D.Math({ children: ommlElements })
```

**第一层：KaTeX 解析 LaTeX → MathML 字符串**

```js
const html = katex.renderToString(latex, { output: 'mathml', throwOnError: false });
const div = document.createElement('div');
div.innerHTML = html;
const mathEl = div.querySelector('math');
```

> **关键**：KaTeX 的 MathML 输出结构是 `math > semantics > mrow > [内容]`，
> `semantics` 下还有 `annotation` 节点（含原始 LaTeX 字符串），需要跳过。

**第二层：MathML → OMML 递归转换（核心映射表）**

| MathML 标签 | docx 对象 | 说明 |
|---|---|---|
| `math`, `mstyle`, `mpadded`, `merror` | 直接展开子节点 | 容器节点 |
| `semantics` | `mathmlNodeToOmml(children[0])` | **必须用 mathmlNodeToOmml 而非 childOmml** |
| `annotation`, `annotation-xml` | 忽略（返回 `[]`） | 含原始 LaTeX，不需要 |
| `mrow` | `tryFencedMrow` 或展开子节点 | 见下文定界符小节 |
| `mi`, `mn`, `mo`, `mtext` | `new D.MathRun(text)` | 原子文本 |
| `mspace` | `new D.MathRun('\u00a0')` | 空白 |
| `mfrac` | `new D.MathFraction({ numerator, denominator })` | 分数 |
| `msqrt` | `new D.MathRadical({ children })` | 平方根 |
| `mroot` | `new D.MathRadical({ children, degree })` | n次根 |
| `msub` | `new D.MathSubScript({ children, subScript })` | 下标 |
| `msup` | `new D.MathSuperScript({ children, superScript })` | 上标 |
| `msubsup` | `new D.MathSubSuperScript({ children, subScript, superScript })` | 上下标 |
| `mover` (accent=true) | `buildAccent(accentChar, base)` | 见下文重音小节 |
| `mover` (accent≠true) | `MathSuperScript` | 极限上标等 |
| `munder` | `MathSubScript` | 极限下标等 |
| `munderover` | `MathSubSuperScript` | 求和上下标 |
| `mtable` | 按列展开，见下文多行公式小节 | 矩阵/cases |
| 其他 | `childOmml(node)` | 安全展开 |

**定界符 `\left...\right`（tryFencedMrow）**

KaTeX 将 `\left( ... \right)` 渲染为：
```xml
<mrow>
  <mo fence="true">(</mo>
  ...内容...
  <mo fence="true">)</mo>
</mrow>
```

检测方式：`mrow` 首尾元素都是 `fence="true"` 的 `<mo>`。

```js
function isFenceMo(el) {
  const tag = el.tagName.toLowerCase().replace(/^[a-z0-9]+:/, '');
  return tag === 'mo' && el.getAttribute('fence') === 'true';
}
```

`\middle|` 产生中间的 `fence="true"` 节点，成为多段分隔符。

生成的 OMML：`<m:d>` 定界符元素，通过 `XmlComponent` 手动构建：
```js
// 简单括号可复用 docx.js 内置类：
new D.MathRoundBrackets({ children })   // ( )
new D.MathSquareBrackets({ children })  // [ ]
new D.MathCurlyBrackets({ children })   // { }

// 自定义定界符（角括号、绝对值、范数等）用 XmlComponent：
const d   = new D.XmlComponent('m:d');
const dPr = new D.XmlComponent('m:dPr');
// 设置 m:begChr, m:sepChr, m:endChr
// 注意：OMML 默认值是 ( | )，相应的默认值可以省略属性
```

**重音符号（buildAccent）**

KaTeX 将 `\widetilde{x}` 渲染为：
```xml
<mover accent="true">
  <mi>x</mi>
  <mo>~</mo>
</mover>
```

需要将 KaTeX 输出的间距字符映射到 OMML 组合字符（Unicode combining characters）：

```js
const ACCENT_MAP = {
  '\u007e': '\u0303', // ~ tilde           → U+0303
  '\u005e': '\u0302', // ^ hat/circumflex  → U+0302 (OMML 默认，可省略 m:chr)
  '\u02c9': '\u0304', // ˉ macron          → U+0304
  '\u00af': '\u0304', // ¯ macron          → U+0304
  '\u0305': '\u0305', // combining overline → keep
  '\u02d9': '\u0307', // ˙ dot above       → U+0307
  '\u00a8': '\u0308', // ¨ diaeresis       → U+0308
  '\u0060': '\u0300', // ` grave           → U+0300
  '\u00b4': '\u0301', // ´ acute           → U+0301
  '\u02d8': '\u0306', // ˘ breve           → U+0306
  '\u02c7': '\u030c', // ˇ caron           → U+030C
  '\u2192': '\u20d7', // → right arrow     → U+20D7
  '\u20d7': '\u20d7', // combining arrow   → keep
};
```

OMML `<m:acc>` 构建：
```js
const acc   = new D.XmlComponent('m:acc');
const accPr = new D.XmlComponent('m:accPr');
// U+0302 是 OMML 默认重音（^ hat），省略 m:chr 让 Word 使用默认
if (ommlChar !== '\u0302') {
  accPr.root.push(new D.MathAccentCharacter(ommlChar));
}
if (accPr.root.length) acc.root.push(accPr);
acc.root.push(new D.MathBase(baseChildren));
```

#### 2.3.3 多行公式（m:eqArr / m:m）

LaTeX 的 `\begin{cases}` / `\begin{aligned}` / `\begin{align}` 等多行环境，
KaTeX 将其渲染为 `<mtable>` 结构（MathML 表格）。
若直接把各行拼成一行，Word 导出后公式失去换行结构，排版混乱。

**OMML 多行元素**

| OMML 元素 | 用途 | 对应 LaTeX 环境 |
|---|---|---|
| `<m:eqArr>` | 方程数组（每行独立换行） | `cases`, `aligned`, `align`, `gather` |
| `<m:m>` | 矩阵（行+列） | `pmatrix`, `bmatrix`, `vmatrix`, etc. |

eqArr 结构：
```xml
<m:eqArr>
  <m:e><!-- 第1行内容 --></m:e>
  <m:e><!-- 第2行内容 --></m:e>
</m:eqArr>
```

matrix 结构：
```xml
<m:m>
  <m:mr>
    <m:e><!-- cell(0,0) --></m:e>
    <m:e><!-- cell(0,1) --></m:e>
  </m:mr>
</m:m>
```

**构建函数**

```js
/** 方程数组：rowContents[i] = 第i行的 OMML 元素数组 */
function buildEqArr(rowContents) {
  const eqArr = new D.XmlComponent('m:eqArr');
  for (const rowContent of rowContents) {
    const e = new D.XmlComponent('m:e');
    e.root.push(...rowContent);
    eqArr.root.push(e);
  }
  return eqArr;
}

/** 矩阵：rows2d[i][j] = 第i行第j列的 OMML 元素数组 */
function buildMatrix(rows2d) {
  const m = new D.XmlComponent('m:m');
  for (const cells of rows2d) {
    const mr = new D.XmlComponent('m:mr');
    for (const cell of cells) {
      const e = new D.XmlComponent('m:e');
      e.root.push(...cell);
      mr.root.push(e);
    }
    m.root.push(mr);
  }
  return m;
}
```

**分发逻辑**

情形1：`<mtable>` 在 fence mrow 内（由 `tryFencedMrow` 处理）

```js
const innerEls = innerNodes.filter(n => n.nodeType === Node.ELEMENT_NODE);
if (innerEls.length === 1 && innerTag === 'mtable') {
  const rows2d = /* 将 mtd 内容转为 OMML 二维数组 */;
  const maxCols = Math.max(...rows2d.map(r => r.length));
  const isMatrixFence = '(['.includes(begChar) && '(['.includes(endChar);

  if (maxCols > 1 && isMatrixFence) {
    // pmatrix / bmatrix 等多列矩阵
    return mkDelimiter(begChar, null, endChar, [[buildMatrix(rows2d)]]);
  } else {
    // cases（单列，begChar='{'）或其他单列有界环境
    const rowContents = rows2d.map(r => r.flat());
    return mkDelimiter(begChar, null, endChar, [[buildEqArr(rowContents)]]);
  }
}
```

情形2：独立 `<mtable>`（`mathmlNodeToOmml` 的 `mtable` case）

```js
case 'mtable': {
  // align / aligned / gather / array 等无 fence 的多行环境
  const rows2d = Array.from(node.children).map(mtr =>
    Array.from(mtr.children).map(mtd => {
      const out = [];
      for (const c of mtd.childNodes) out.push(...mathmlNodeToOmml(c));
      return out;
    })
  );
  if (rows2d.length === 0) return [];
  // 多列 align 的左右对齐列合并为一行（x &= y → [x,=,y]）
  const rowContents = rows2d.map(r => r.flat());
  return [buildEqArr(rowContents)];
}
```

cases 端到端示例：

LaTeX：`\begin{cases} q^* = \varphi(q^*) \\ Aq^* = 0 \end{cases}`

```xml
<!-- KaTeX MathML（简化） -->
<mrow>
  <mo fence="true">{</mo>
  <mtable>
    <mtr><mtd><mrow>q* = φ(q*)</mrow></mtd></mtr>
    <mtr><mtd><mrow>Aq* = 0</mrow></mtd></mtr>
  </mtable>
  <mo fence="true"></mo>   <!-- 空字符结尾 = 无右括号 -->
</mrow>
```

处理流程：
1. `mathmlNodeToOmml(mrow)` → `tryFencedMrow`
2. begChar=`{`，endChar=`""`（空），无中间 fence
3. innerEls=[mtable]，单列 → `buildEqArr` + `mkDelimiter('{', null, '', [...])`
4. OMML：`<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/></m:dPr><m:e><m:eqArr>...</m:eqArr></m:e></m:d>`

Word 最终效果：
```
⎧ q* = φ(q*)
⎨
⎩ Aq* = 0
```

> 注意：`endChar = ""` 对应 OMML `<m:endChr m:val=""/>` → Word 显示无右括号（cases 正确行为）。
> 嵌套多行（如 `\begin{cases}` 内含 `\begin{aligned}`）：递归调用 `mathmlNodeToOmml` 自然处理。

### 2.4 图片块

#### 2.4.1 数据结构

```js
{
  type: 'img',
  src: 'data:image/jpeg;base64,/9j/4AAQ…',  // base64 data URL（JPEG / PNG）
  caption: '图1. 仿真轨迹示意图。',           // 可选图注
  widthPt: 440,    // 展示宽度（pt），用于 HTML max-width 和 docx ImageRun
  heightPt: 329,   // 对应高度（pt），保持纵横比
}
```

> `src` 推荐使用 `data:image/jpeg;base64,…` 格式，完全内嵌在 HTML 文件中，无需外部依赖。
> 使用 Python `base64.b64encode(raw).decode()` 生成，再拼接前缀即可。

#### 2.4.2 HTML 渲染（renderContent）

```js
case 'img': {
  const cap  = item.caption ? `<figcaption>${esc(item.caption)}</figcaption>` : '';
  const style = item.widthPt ? ` style="max-width:${item.widthPt}pt"` : '';
  html += `<figure class="df-figure">
    <img src="${item.src}"${style} alt="${escAttr(item.caption || '')}">
    ${cap}
  </figure>`;
  break;
}
```

CSS 样式：

```css
.df-figure {
  display: block; margin: 14pt auto; text-align: center;
}
.df-figure img {
  max-width: 100%; height: auto; display: block; margin: 0 auto;
}
.df-figure figcaption {
  font-family: var(--body-font); font-size: var(--body-size);
  line-height: var(--body-lh); color: #333;
  margin-top: 5pt; text-align: center;
}
```

#### 2.4.3 DOCX 导出（exportToDocx）

```js
/** image src（data URL / base64 / http(s) URL）→ Uint8Array */
async function b64ToUint8(src) {
  if (!src) return null;
  const isDataUrl = src.startsWith('data:');
  const isUrl = /^[a-z][a-z0-9+.-]*:/i.test(src) || src.startsWith('//');
  if (isDataUrl || !isUrl) {
    const b64 = src.includes(',') ? src.split(',')[1] : src;
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }
  const response = await fetch(src);
  if (!response.ok) throw new Error(`Image fetch failed for ${src}: ${response.status}`);
  const buf = await response.arrayBuffer();
  return new Uint8Array(buf);
}

async function buildImagePara(src, widthPt, heightPt, caption) {
  const imageRun = new D.ImageRun({
    data: await b64ToUint8(src),
    transformation: { width: widthPt, height: heightPt }
  });
  const imgPara = new D.Paragraph({
    children: [imageRun],
    alignment: D.AlignmentType.CENTER,
    spacing: { before: 120, after: caption ? 60 : 180 }
  });
  if (!caption) return [imgPara];
  const capPara = new D.Paragraph({
    children: [new D.TextRun({ text: caption, font: bodyFontName, size: bodySize, italics: true })],
    alignment: D.AlignmentType.CENTER,
    spacing: { after: 180 }
  });
  return [imgPara, capPara];
}
```

在 `switch(item.type)` 中：

```js
case 'img': {
  const wPt = item.widthPt || 400;
  const hPt = item.heightPt || Math.round(wPt * 0.75);
  docChildren.push(...await buildImagePara(item.src, wPt, hPt, item.caption || ''));
  break;
}
```

> docx.js 版本注意：`ImageRun` 的 `transformation` 在 7.x 和 8.x 中均支持 pt 单位（内部自动转换为 EMU：1pt = 12700 EMU）。

#### 2.4.4 图片来源：从 PDF 提取（Python）

```python
import fitz   # pip install pymupdf
import base64
from PIL import Image
from io import BytesIO

def extract_and_encode(pdf_path, page_idx, clip_rect, scale=2.2, quality=85):
    """提取 PDF 特定页面的裁切区域，返回 base64 JPEG 字符串和尺寸。"""
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    clip = fitz.Rect(*clip_rect)       # (x0, y0, x1, y1) in points
    mat  = fitz.Matrix(scale, scale)   # 放大倍率
    pix  = page.get_pixmap(matrix=mat, clip=clip)

    img = Image.open(BytesIO(pix.tobytes('png'))).convert('RGB')
    if img.width > 1200:               # 限制最大宽度节省体积
        h = round(img.height * 1200 / img.width)
        img = img.resize((1200, h), Image.LANCZOS)

    buf = BytesIO()
    img.save(buf, 'JPEG', quality=quality)
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f'data:image/jpeg;base64,{b64}', img.size

# 示例：提取论文第5页（索引4）上半部分的仿真轨迹图
src, (w, h) = extract_and_encode(
    'paper.pdf', page_idx=4,
    clip_rect=(0, 10, 612, 468),       # PDF坐标（pt），A4宽612pt
    scale=2.2, quality=85
)
# clip_rect 使用 PDF 点坐标（1pt = 1/72 inch）。
# 可先用 page.rect 获取页面尺寸，再在 PDF 查看器中测量目标区域，
# 或先导出整页图像后在图像编辑器中读出像素范围，再按比例换算为 pt。
target_w = 440
target_h = round(target_w * h / w)
entry = f"{{type:'img',src:'{src}',caption:'图1. ...',widthPt:{target_w},heightPt:{target_h}}}"
```

#### 2.4.5 图片来源：从 DOCX 提取（Python）

```python
import zipfile, base64
from PIL import Image
from io import BytesIO

def extract_from_docx(docx_path, max_w=1100, quality=82):
    """从 .docx 文件中提取所有嵌入图片，返回 [(data_url, (w, h)), ...] 列表。"""
    results = []
    with zipfile.ZipFile(docx_path) as z:
        media = sorted(n for n in z.namelist() if n.startswith('word/media/'))
        for name in media:
            raw = z.read(name)
            img = Image.open(BytesIO(raw)).convert('RGB')
            if img.width > max_w:
                h = round(img.height * max_w / img.width)
                img = img.resize((max_w, h), Image.LANCZOS)
            buf = BytesIO()
            img.save(buf, 'JPEG', quality=quality)
            b64 = base64.b64encode(buf.getvalue()).decode()
            results.append((f'data:image/jpeg;base64,{b64}', img.size))
    return results
```

---

## 三、Word 导出（docx.js）

### 3.1 核心 API 速查

**单位换算（必记）**

| CSS 单位 | docx 单位 | 换算 |
|---|---|---|
| pt | half-point | `Math.round(pt * 2)` |
| cm | twip | `cm * 567` (1cm = 566.93 twips) |
| pt（行距） | twip | `pt * 20` |
| 字符宽度（首行缩进）| twip | `bodySize_halfpt / 2 * 2 * 20` = `bodySize_halfpt * 20` |

**常用对象**

```js
const D = window.docx;

// 文字片段
new D.TextRun({ text, font: 'SimSun', size: 24, bold: true, italic: false, color: '000000' })

// 段落
new D.Paragraph({
  children: [...],                              // TextRun / Math 数组
  alignment: D.AlignmentType.CENTER,           // LEFT / CENTER / RIGHT / JUSTIFIED
  spacing: { before: 480, after: 240,          // twips
             line: 480, lineRule: D.LineRuleType.AUTO },
  indent: { firstLine: 480, left: 480, hanging: 480 }
})

// 数学公式（OMML 包装器）
new D.Math({ children: ommlElements })

// 文档
new D.Document({
  styles: { paragraphStyles: [] },
  sections: [{
    properties: { page: { margin: { top, bottom, left, right, gutter } } },
    headers: { default: new D.Header({ children: [...] }) },
    children: docChildren
  }]
})

// 导出
const blob = await D.Packer.toBlob(doc);
saveAs(blob, 'file.docx');
```

**页面尺寸（A4，SEU 格式示例）**

```js
margin: {
  top:    1134,  // 2 cm
  bottom: 1134,  // 2 cm
  left:   1417,  // 2.5 cm (含装订线)
  right:  1134,  // 2 cm
  gutter: 284    // 0.5 cm 装订线
}
```

### 3.2 分页、页眉与页码

**功能目标**

- 页码支持 Word/PDF 同步配置：
  - 位置：不显示 / 页眉居左 / 页眉居中 / 页眉居右 / 页脚居左 / 页脚居中 / 页脚居右
  - 格式：`n`、`第 n 页`、`n / N`、`第 n 页 / 共 N 页`
- 页眉文本支持自定义内容和左/中/右位置
- 页眉文本与页码共存于同一行（不同位置），通过 **tab stop** 机制实现

**UI 配置**

```html
<input id="s-header-text">
<select id="s-header-align">   <!-- left / center / right / none -->
<select id="s-page-num-position">  <!-- none / header-left / header-center / header-right / footer-* -->
<select id="s-page-num-format">    <!-- n / nN / cn / cnN -->
```

```js
function getPageDecorSettings() {
  return {
    headerText:      document.getElementById('s-header-text').value,
    headerAlign:     document.getElementById('s-header-align').value,    // 'left'|'center'|'right'|'none'
    pageNumPosition: document.getElementById('s-page-num-position').value,
    pageNumFormat:   document.getElementById('s-page-num-format').value
  };
}
```

**docx.js 实现：用 tab stop 把文本和页码放在同一行**

Word 页眉的标准做法是在同一段落里用 tab stop 分隔左/中/右内容。
不能用两个独立段落——那会造成页眉占两行，且 UI 选择的对齐位置完全无效。

A4 正文宽度（对应 tab stop 端点）：
```
正文宽 = 页面宽(11906 twips) - 左边距(1417) - 右边距(1134) = 9355 twips
```

```js
const HDR_TEXT_WIDTH = 9355;                      // A4 正文宽（twips）
const hdrTabStops = [
  { type: D.TabStopType.CENTER, position: Math.round(HDR_TEXT_WIDTH / 2) },  // 4677 twips
  { type: D.TabStopType.RIGHT,  position: HDR_TEXT_WIDTH }                    // 9355 twips
];

// 按位置把页眉文本和页码分别填入三个槽位
const hdrL = [], hdrC = [], hdrR = [];
const pnPos = pageDecor.pageNumPosition;
if (pageDecor.headerAlign !== 'none') {
  const t = new D.TextRun({ text: pageDecor.headerText, font: bodyFontName, size: 18 });
  if      (pageDecor.headerAlign === 'left')   hdrL.push(t);
  else if (pageDecor.headerAlign === 'center') hdrC.push(t);
  else if (pageDecor.headerAlign === 'right')  hdrR.push(t);
}
if (pnPos === 'header-left')        hdrL.push(...docxPageRuns());
else if (pnPos === 'header-center') hdrC.push(...docxPageRuns());
else if (pnPos === 'header-right')  hdrR.push(...docxPageRuns());

const headers = {
  default: new D.Header({
    children: [new D.Paragraph({
      tabStops: hdrTabStops,
      children: [
        ...hdrL,
        new D.TextRun({ text: '\t' }),   // 跳到居中位置
        ...hdrC,
        new D.TextRun({ text: '\t' }),   // 跳到居右位置
        ...hdrR,
      ],
      border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } }
    })]
  })
};

// 页脚：同样用 tab stop 定位，仅放置页码
let footers;
if (pnPos.startsWith('footer-')) {
  const ftrRuns = pnPos === 'footer-left'
    ? [...docxPageRuns()]
    : pnPos === 'footer-center'
      ? [new D.TextRun({ text: '\t' }), ...docxPageRuns()]
      : [new D.TextRun({ text: '\t' }), new D.TextRun({ text: '\t' }), ...docxPageRuns()];
  footers = { default: new D.Footer({ children: [
    new D.Paragraph({ tabStops: hdrTabStops, children: ftrRuns })
  ]})};
}
```

页码动态字段（支持总页数）：

```js
function docxPageRuns() {
  if (!(D.PageNumber && D.PageNumber.CURRENT)) return [new D.TextRun({ text: '1', font: bodyFontName, size: 18 })];
  const hasTotalPage = !!(D.PageNumber && D.PageNumber.TOTAL_PAGES);
  switch (pageDecor.pageNumFormat) {
    case 'n':   return [new D.TextRun({ children: [D.PageNumber.CURRENT], font: bodyFontName, size: 18 })];
    case 'nN':  return hasTotalPage
      ? [new D.TextRun({ children: [D.PageNumber.CURRENT, ' / ', D.PageNumber.TOTAL_PAGES], font: bodyFontName, size: 18 })]
      : [new D.TextRun({ children: [D.PageNumber.CURRENT], font: bodyFontName, size: 18 })];
    case 'cn':  return [new D.TextRun({ children: ['第 ', D.PageNumber.CURRENT, ' 页'], font: bodyFontName, size: 18 })];
    case 'cnN':
    default:    return hasTotalPage
      ? [new D.TextRun({ children: ['第 ', D.PageNumber.CURRENT, ' 页 / 共 ', D.PageNumber.TOTAL_PAGES, ' 页'], font: bodyFontName, size: 18 })]
      : [new D.TextRun({ children: ['第 ', D.PageNumber.CURRENT, ' 页'], font: bodyFontName, size: 18 })];
  }
}
```

**已修复的 Bug（页眉页脚不导出问题）**

| # | 现象 | 根因 | 修复 |
|---|---|---|---|
| 1 | 页眉文本和页码在 Word 中占两行 | `headerChildren` 用两个独立段落而非 tab stop 同行布局 | 改为单段落 + tab stop，文本和页码共行 |
| 2 | 选择"页眉居左/居右"页码后，导出位置始终在右 | `pageNumAlign` 用 `endsWith('center')` 判断，`'header-left'` 映射到 RIGHT | 用 tab stop 槽位分配，彻底去掉 `pageNumAlign` |
| 3 | `headerAlign = 'none'` 时页眉文本仍然出现 | 代码直接 `pageDecor.headerText` 放进段落，未检查 `headerAlign` | 先检查 `headerAlign !== 'none'` 再填入槽位 |

---

## 四、PDF 导出（html2pdf.js）

### 4.1 基础用法

```js
html2pdf().set({
  margin: 0,
  filename: 'report.pdf',
  image: { type: 'jpeg', quality: 0.97 },
  html2canvas: { scale: 2, useCORS: true },  // scale:2 保证清晰度
  jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
}).from(document.getElementById('document-page')).save();
```

> **优点**：数学公式由 KaTeX 渲染好再截图，字体完全与网页一致。
> **缺点**：无法矢量化文字，文字层不可选中；中文字体需依赖系统字体。

### 4.2 分页与页码实现

- 不直接导出整页长 DOM；改为构建临时容器，每个 `.pdf-export-page` 固定 A4 高度（297mm）
- 内容区逐块填充：超高即新建下一页；遇到显式分页点（如"参考文献"）强制换页
- 生成完总页数后，回填每页页码文本，再调用 `html2pdf`

```js
let cur = makePage();
for (const block of Array.from(sourceContent.children)) {
  const clone = block.cloneNode(true);
  // 强制换页逻辑
  if (clone.classList?.contains('dh1') &&
      clone.textContent.trim() === '参考文献' &&
      cur.contentInner.children.length > 0) {
    cur = makePage();
  }
  cur.contentInner.appendChild(clone);
  if (cur.contentOuter.scrollHeight > cur.contentOuter.clientHeight + 2) {
    cur.contentInner.removeChild(clone);
    cur = makePage();
    cur.contentInner.appendChild(clone);
  }
}
```

### 4.3 空白页问题排查与规范

PDF 导出中出现多余空白页有两类来源，需分别处理：

#### A. 每页后都有空白页（双重断页）

原因：`.pdf-export-page` 设置了 `page-break-after: always`，同时 html2pdf 选项 `pagebreak: { mode: ['css', 'legacy'] }` 会扫描该属性并调用 `jsPDF.addPage()`，与自然按高度切页叠加，导致每页内容后插入一张空白页。

**修复方法**：禁用 html2pdf 的 CSS 断页扫描：
```js
pagebreak: { mode: [] }
```
此时每个 297mm 高度的 div 自然对应一页，无需额外断页指令。

#### B. 最后一页多一张空白页（像素取整溢出）

原因：CSS `height: 297mm` 在浏览器中转为像素时向上取整（297/25.4×96 = 1122.52 → 1123px），而 html2pdf 内部按 `Math.floor(canvasWidth × 297/210)` 计算每页像素高度。两者不完全整除时，`Math.ceil(totalHeight / pageHeight)` 会多出 1 页空白。

**修复方法**：将 `html2canvas.height` 设为与 html2pdf 内部 `pxPageHeight` 一致的整数倍：
```js
html2canvas: {
  width:  container.offsetWidth  || 794,
  height: Math.floor((container.offsetWidth || 794) * 297 / 210) * pages.length
}
```
这样 `canvas.height / pxPageHeight = pages.length`，整除无余，不生成额外页。

### 4.4 兼容性建议

- `TOTAL_PAGES` 不可用时，自动降级为仅显示当前页码
- 对公式块、图块设置 `break-inside: avoid`，减少跨页切断

---

## 五、参考与调试

### 5.1 MathML → OMML 覆盖情况

| LaTeX 类别 | KaTeX MathML 标签 | OMML 处理 | 状态 |
|---|---|---|---|
| 分数 `\frac` | `<mfrac>` | `MathFraction` | ✅ |
| 平方根 `\sqrt` | `<msqrt>` | `MathRadical` | ✅ |
| n次根 `\sqrt[n]` | `<mroot>` | `MathRadical` with degree | ✅ |
| 下标 `_` | `<msub>` | `MathSubScript` | ✅ |
| 上标 `^` | `<msup>` | `MathSuperScript` | ✅ |
| 上下标 `_^` | `<msubsup>` | `MathSubSuperScript` | ✅ |
| 重音 `\hat\bar\vec` 等 | `<mover accent="true">` | `m:acc` via XmlComponent | ✅ |
| 极限上标 `\overset` | `<mover accent≠true>` | `MathSuperScript`（降级）| ⚠️ |
| 极限下标 `\underset` | `<munder>` | `MathSubScript`（降级）| ⚠️ |
| 定积分上下限 `\munderover` | `<munderover>` | `MathSubSuperScript` | ✅ |
| `\left...\right` 定界符 | `<mrow>` with fence mo | `m:d` via XmlComponent | ✅ |
| `\middle` | fence mo 中间节点 | 多段 `m:d` | ✅ |
| 矩阵 `pmatrix/bmatrix` | `<mrow>+<mtable>` (多列) | `m:m` via XmlComponent（在 tryFencedMrow 检测）| ✅ |
| cases 环境 | `<mrow fence={>+<mtable>` (单列) | `m:eqArr` via XmlComponent（换行分行）| ✅ |
| align/aligned 环境 | 独立 `<mtable>` | `m:eqArr`（多列单元格合并后换行）| ✅ |
| 文本 `\text{}` | `<mtext>` | `MathRun` | ✅ |
| 空白 `\;` 等 | `<mspace>` | `MathRun('\u00a0')` | ✅ |
| 希腊字母 | `<mi>` | `MathRun`（直接 Unicode）| ✅ |
| 运算符 `+−×÷` 等 | `<mo>` | `MathRun` | ✅ |

> ⚠️ = 功能性降级，输出可读但非完美排版。

### 5.2 已知 Bug 与修复

**Bug 1：所有 `\left...\right` 定界符渲染成三个散字符**

**现象**：`\left(x\right)` 在 Word 中显示为 `(x)` 而非正确缩放的括号。

**根因**：`semantics` 分支调用了 `childOmml(node.children[0])` 而非 `mathmlNodeToOmml(node.children[0])`。
- `childOmml(mrow)` = 直接展开 mrow 的子节点，跳过 mrow 自身
- KaTeX 输出始终是 `math > semantics > mrow > 内容`
- 因此 `tryFencedMrow` 永远不被调用

**修复**：
```js
// ❌ 错误 - 跳过了 mrow 的处理
case 'semantics':
  return node.children[0] ? childOmml(node.children[0]) : [];

// ✅ 正确 - mrow 会走到自己的 case，tryFencedMrow 被调用
case 'semantics':
  return node.children[0] ? mathmlNodeToOmml(node.children[0]) : [];
```

**Bug 2：重音符号（\widetilde 等）渲染成上标**

**现象**：`\widetilde{W}` 在 Word 中显示为 `W~` 而非 $\widetilde{W}$。

**根因**：`mover` 分支未检查 `accent="true"` 属性，统一用 `MathSuperScript` 处理。

**修复**：增加 `accent="true"` 判断，走 `buildAccent` 路径。

### 5.3 开发调试技巧

1. **在浏览器控制台测试单个 LaTeX**：
   ```js
   katex.renderToString('\\widetilde{W}', {output:'mathml', throwOnError:false});
   ```

2. **检查 MathML 结构**：
   ```js
   const div = document.createElement('div');
   div.innerHTML = katex.renderToString('\\left(x\\right)', {output:'mathml'});
   console.log(div.querySelector('math').outerHTML);
   ```

3. **验证 `fence="true"` 是否正确检测**：
   观察 `<mrow>` 的首尾 `<mo>` 是否带 `fence="true"` 属性。
   - `\left...\right` → 会有 `fence="true"`
   - `\{` (不带 left/right) → `stretchy=false`，**没有** `fence="true"`

4. **docx.js 对象构建失败的通用降级**：
   ```js
   try {
     return [new D.MathRadical({ children: ... })];
   } catch(e) {
     return [new D.MathRun('√'), ...children];  // 纯文本降级
   }
   ```

5. **Node.js 环境模拟测试**（不需要浏览器）：
   - 安装 `katex` + `jsdom`
   - 用 jsdom 提供 DOM API，把完整的 `mathmlNodeToOmml` 逻辑复制进去测试
   - 用 `check(tree)` lambda 断言输出结构

### 5.4 可直接复用的最简模板

```html
<!DOCTYPE html>
<html>
<head>
<!-- KaTeX -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
<!-- docx.js + FileSaver -->
<script src="https://unpkg.com/docx@8/build/index.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
</head>
<body>
<button onclick="exportDocx()">下载 Word</button>
<script>
async function exportDocx() {
  // 1. 等待库加载
  await new Promise(res => {
    const iv = setInterval(() => {
      if (window.docx && typeof saveAs === 'function') { clearInterval(iv); res(); }
    }, 200);
  });

  const D = window.docx;

  // 2. LaTeX → OMML（复制 mathmlNodeToOmml + latexToOmml 两个函数）
  // ...（见完整实现）

  // 3. 构建文档
  const doc = new D.Document({
    sections: [{
      children: [
        new D.Paragraph({ children: [new D.TextRun({ text: 'Hello World', size: 24 })] }),
        new D.Paragraph({ children: [new D.Math({ children: latexToOmml('\\frac{a}{b}') })] })
      ]
    }]
  });

  // 4. 下载
  const blob = await D.Packer.toBlob(doc);
  saveAs(blob, 'output.docx');
}
</script>
</body>
</html>
```

---

## 六、UI 界面层（report_editor.html）

本节描述与导出功能配套的完整单页文档编辑器 UI，包括整体布局、侧边栏控件、样式模板、实时预览机制和通知系统。

### 6.1 整体页面布局

```
┌─────────────────────────────────────────────────────┐
│  #topbar（46px 顶栏）                                │
│  背景 #1a3a6b（深蓝），标题 + 副标题                 │
├────────────────┬────────────────────────────────────┤
│  #sidebar      │  #main                             │
│  (272px 固定)  │  (flex:1, overflow-y:auto)         │
│                │                                    │
│  左侧样式面板  │  居中显示 A4 文档页（#document-page）│
│  可纵向滚动    │  背景 #dde1e8                       │
└────────────────┴────────────────────────────────────┘
```

关键布局 CSS：
```css
body { overflow: hidden; }         /* 禁止 body 滚动，由子区域各自滚动 */
#app  { display: flex; height: calc(100vh - 46px); }
#sidebar { width: 272px; overflow-y: auto; display: flex; flex-direction: column; }
#main    { flex: 1; overflow-y: auto; padding: 20px 28px; }
```

### 6.2 A4 文档预览区（#document-page）

```css
#document-page {
  width: 210mm; min-height: 297mm;
  margin: 0 auto;
  padding: 2cm 2cm 2cm 2.5cm;   /* 上下左右：SEU 格式，左装订线 0.5cm 额外 */
  background: #fff;
  box-shadow: 0 3px 14px rgba(0,0,0,.2);

  /* === CSS 变量（由侧边栏控件实时更新）=== */
  --body-font: 'SimSun','宋体',serif;
  --body-size: 12pt;
  --body-lh: 1.5;
  --h1-font: 'SimHei','黑体',sans-serif;  --h1-size: 15pt;  --h1-align: center;
  --h2-font: 'SimHei','黑体',sans-serif;  --h2-size: 14pt;  --h2-align: left;
  --h3-font: 'SimSun','宋体',serif;       --h3-size: 12pt;  --h3-align: left;
}
```

页眉预览（`.pg-header`）：三个 `<span>`（left / center / right）用 flexbox 分布，由 `applyPageDecorSettings()` 填充文本。

### 6.3 侧边栏（#sidebar）各区块

侧边栏由多个 `.sb-section` 堆叠，每区块有 `.sb-title`（大写灰色标签）和若干控件。

#### 区块1：📐 样式模板

```html
<div class="tpl-wrap" id="tpl-wrap">
  <div class="tpl active" data-tpl="seu"      onclick="loadTemplate(this)">SEU标准</div>
  <div class="tpl"        data-tpl="academic" onclick="loadTemplate(this)">学术简洁</div>
  <div class="tpl"        data-tpl="modern"   onclick="loadTemplate(this)">现代简约</div>
  <div class="tpl"        data-tpl="classic"  onclick="loadTemplate(this)">经典学报</div>
</div>
```

四套预设模板（`TEMPLATES` 对象）：

| 模板 key | 正文字体 | 正文字号 | 行距 | H1字体 | H1字号 | H1对齐 |
|---|---|---|---|---|---|---|
| `seu`      | SimSun 宋体     | 12pt  | 1.5 | SimHei 黑体 | 15pt | 居中 |
| `academic` | Times New Roman | 12pt  | 1.6 | Arial       | 16pt | 居中 |
| `modern`   | Arial           | 12pt  | 1.7 | Arial       | 18pt | 左对齐 |
| `classic`  | SimSun 宋体     | 10.5pt| 1.5 | SimHei 黑体 | 16pt | 居中 |

点击芯片 → `loadTemplate(el)` → 批量调用 `setSelect(id, val)` 设置所有 `<select>` → `applyStyles()`。

`setSelect` 的容错设计：若目标 `<option>` 不存在则动态创建后选中，不报错。

#### 区块2：📝 正文样式

| 控件 ID | 类型 | 作用 | 默认值 |
|---|---|---|---|
| `s-bfont`  | `<select>` | 正文字体 | `'SimSun','宋体',serif` |
| `s-bsize`  | `<select>` | 正文字号 | `12pt`（小四）|
| `s-blh`    | `<select>` | 行距倍数 | `1.5` |

`s-bsize` 和 `s-blh` 通过 `.grid2`（`grid-template-columns: 1fr 1fr`）并排显示。

#### 区块3-5：📌 标题样式（一/二/三级）

| 控件 ID | 类型 | 默认值 |
|---|---|---|
| `s-h1font` | `<select>` 字体 | SimHei 黑体 |
| `s-h1size` | `<select>` 字号 | 15pt（小三）|
| `s-h1align`| `<select>` 对齐 | 居中 |
| `s-h2font` | `<select>` 字体 | SimHei 黑体 |
| `s-h2size` | `<select>` 字号 | 14pt（四号）|
| `s-h2align`| `<select>` 对齐 | 左对齐 |
| `s-h3size` | `<select>` 字号 | 12pt（小四）|
| `s-h3align`| `<select>` 对齐 | 左对齐 |

三级标题没有独立字体选项（继承正文字体或默认）。

每个控件均绑定 `onchange="applyStyles()"` 触发实时预览更新。

#### 区块4：📄 页眉与页码

| 控件 ID | 类型 | 选项 | 作用 |
|---|---|---|---|
| `s-header-text`     | `<input>`  | — | 页眉文字内容 |
| `s-header-align`    | `<select>` | `left`/`center`/`right`/`none` | 页眉文字位置 |
| `s-page-num-position`| `<select>`| `none` / `header-left` / `header-center` / `header-right` / `footer-left` / `footer-center` / `footer-right` | 页码位置 |
| `s-page-num-format` | `<select>` | `n` / `nN` / `cn` / `cnN` | 页码格式 |

**冲突防御**（`updatePageDecorConflicts()`）：
- 页眉文字已占用某位置时，页码位置选项中对应项自动 `disabled`
- 页码已占用某位置时，页眉文字对齐选项中对应项自动 `disabled`
- 避免同一位置同时显示文字和页码

所有控件均绑定 `oninput="onPageDecorChange()"` 或 `onchange="onPageDecorChange()"`。

`onPageDecorChange()` 调用链：
```
onPageDecorChange()
  └─ updatePageDecorConflicts()   // 禁用冲突选项
  └─ applyPageDecorSettings()     // 更新 #document-page 中 .pg-header 的三个 span
```

#### 区块5：下载按钮区（`.dl-area`）

```html
<div class="dl-area">  <!-- margin-top: auto → 始终固定在侧边栏底部 -->
  <button class="btn btn-docx" onclick="exportToDocx()">⬇ 下载 DOCX（含Word公式）</button>
  <button class="btn btn-pdf"  onclick="exportToPDF()">⬇ 下载 PDF</button>
</div>
```

| 按钮类 | 背景色 | 触发函数 |
|---|---|---|
| `.btn-docx` | `#1f4e79`（深蓝）| `exportToDocx()` |
| `.btn-pdf`  | `#c00000`（深红）| `exportToPDF()` |

`.dl-area` 使用 `margin-top: auto` 将按钮推到侧边栏底部（sidebar 为 `flex-column`）。

### 6.4 `applyStyles()` — 实时 CSS 变量注入

```js
function applyStyles() {
  const pg  = document.getElementById('document-page');
  const set = (k, v) => pg.style.setProperty(k, v);

  set('--body-font',  document.getElementById('s-bfont').value);
  set('--body-size',  document.getElementById('s-bsize').value);
  set('--body-lh',    document.getElementById('s-blh').value);
  set('--h1-font',    document.getElementById('s-h1font').value);
  set('--h1-size',    document.getElementById('s-h1size').value);
  set('--h1-align',   document.getElementById('s-h1align').value);
  set('--h2-font',    document.getElementById('s-h2font').value);
  set('--h2-size',    document.getElementById('s-h2size').value);
  set('--h2-align',   document.getElementById('s-h2align').value);
  set('--h3-size',    document.getElementById('s-h3size').value);
  set('--h3-align',   document.getElementById('s-h3align').value);
  onPageDecorChange();   // 同步更新页眉/页码预览
}
```

CSS 变量仅设置在 `#document-page` 节点上，作用域即文档预览区，不影响侧边栏样式。导出时从相同 `<select>` 读取同一套值，保证预览与导出一致。

### 6.5 内容块 DOM 类名与 CONTENT 类型对应

| CONTENT `type` | DOM 类名 | CSS 变量 | contenteditable |
|---|---|---|---|
| `h0` / `h1` | `.dh1` | `--h1-*` | `true` |
| `h2` | `.dh2` | `--h2-*` | `true` |
| `h3` | `.dh3` | `--h3-*` | `true` |
| `para` | `.dp` | `--body-*`，`text-indent:2em` | `true` |
| `kw`（关键词行）| `.dkw` | `--body-*` | `true` |
| `formula` | `.df-wrap > .df-inner` | — | `false`（KaTeX 渲染）|
| `refs` | `.dref` | `--body-*`，悬挂缩进 `text-indent:-2em;padding-left:2em` | `true` |
| `img` | `figure.df-figure` | — | `false` |

公式块额外结构：
```html
<div class="df-wrap">           <!-- flex 容器，相对定位 -->
  <div class="df-inner" data-latex="...">
    <span class="math-block">\[...\]</span>   <!-- KaTeX auto-render 目标 -->
  </div>
  <span class="df-num">(1)</span>              <!-- 公式编号，绝对定位在右 -->
</div>
```

`renderContent()` 在页面初始化时执行一次，之后由 KaTeX `auto-render` 扫描 `.math-block` 并渲染。

### 6.6 通知组件（#notif）

```html
<div id="notif"></div>
```

```css
#notif {
  position: fixed; top: 60px; right: 18px;
  background: rgba(30,30,30,.92); color: #fff;
  padding: 9px 18px; border-radius: 7px;
  opacity: 0; transition: opacity .25s; pointer-events: none;
}
#notif.show { opacity: 1; }
```

```js
let notifTimer = null;
function notify(msg, duration = 2500) {
  const el = document.getElementById('notif');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(notifTimer);
  notifTimer = setTimeout(() => el.classList.remove('show'), duration);
}
```

在导出开始/结束/报错时调用：
```js
notify('正在生成 DOCX，请稍候…');
notify('✅ DOCX 已下载！');
notify('❌ 导出失败：' + err.message, 4000);
```

### 6.7 库加载状态指示（#lib-status）

页面底部有一个固定定位的状态条，实时显示 CDN 库的加载进度：

```js
function updateLibStatus() {
  const docxOk     = !!(window.docx && window.docx.Document);
  const saverOk    = typeof saveAs === 'function';
  const katexOk    = typeof katex  === 'function';
  const pdf2Ok     = typeof html2pdf === 'function';
  // 全部就绪时隐藏状态条，否则显示各库加载状态
}
```

这允许用户在库尚未加载完成时（网络慢）看到明确提示，而非点击导出后静默失败。

### 6.8 完整 HTML 骨架

```html
<!-- 顶栏 -->
<div id="topbar">
  <h1>结项报告编辑器</h1>
  <span class="subtitle">| 项目名称</span>
</div>

<!-- 主体：侧边栏 + 文档预览 -->
<div id="app">
  <aside id="sidebar">
    <!-- 样式模板芯片 -->
    <div class="sb-section"> ... </div>
    <!-- 正文样式 -->
    <div class="sb-section"> ... </div>
    <!-- 一/二/三级标题样式 -->
    <div class="sb-section"> ... </div>
    <div class="sb-section"> ... </div>
    <div class="sb-section"> ... </div>
    <!-- 页眉与页码 -->
    <div class="sb-section"> ... </div>
    <!-- 下载按钮（margin-top:auto 固定底部）-->
    <div class="dl-area">
      <button class="btn btn-docx" onclick="exportToDocx()">⬇ 下载 DOCX</button>
      <button class="btn btn-pdf"  onclick="exportToPDF()">⬇ 下载 PDF</button>
    </div>
  </aside>

  <main id="main">
    <div id="document-page">
      <!-- 页眉预览 -->
      <div class="pg-header">
        <span id="pg-hdr-left"></span>
        <span id="pg-hdr-center"></span>
        <span id="pg-hdr-right">第 1 页</span>
      </div>
      <!-- 正文内容（由 renderContent() 填充）-->
      <div id="doc-content"></div>
    </div>
  </main>
</div>

<!-- 全局通知 -->
<div id="notif"></div>

<!-- 库加载状态 -->
<div id="lib-status"></div>
```
