# SKILL: 从 HTML 网页生成可直接下载的 Word (.docx) 文档

> 适用场景：单页 HTML 应用、学术报告、科技文档，需在浏览器端（无后端）将富文本 + LaTeX 公式导出为规范 Word 文件。

---

## 一、整体架构

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
    │       └─ 独立公式  → buildMathForLatex → Math (OMML)
    ├─ 4. 构建 Document（页面尺寸、页边距、页眉）
    ├─ 5. Packer.toBlob(doc)
    └─ 6. FileSaver.saveAs(blob, 'xxx.docx')
```

---

## 二、关键依赖库（CDN 方式引入）

| 库 | 版本 | 作用 | CDN |
|---|---|---|---|
| **docx.js** | 8.x | 构建 `.docx` 的所有元素（段落、文字、公式、页眉等） | `https://unpkg.com/docx@8/build/index.js` |
| **FileSaver.js** | 2.x | `saveAs(blob, filename)` 触发浏览器下载 | `https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js` |
| **KaTeX** | 0.16.x | LaTeX → MathML 解析（用于后续转 OMML） | `https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js` + auto-render |

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

## 三、docx.js 核心 API 速查

### 3.1 单位换算（必记）

| CSS 单位 | docx 单位 | 换算 |
|---|---|---|
| pt | half-point | `Math.round(pt * 2)` |
| cm | twip | `cm * 567` (1cm = 566.93 twips) |
| pt（行距） | twip | `pt * 20` |
| 字符宽度（首行缩进）| twip | `bodySize_halfpt / 2 * 2 * 20` = `bodySize_halfpt * 20` |

### 3.2 常用对象

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

### 3.3 页面尺寸（A4，SEU 格式示例）

```js
margin: {
  top:    1134,  // 2 cm
  bottom: 1134,  // 2 cm
  left:   1417,  // 2.5 cm (含装订线)
  right:  1134,  // 2 cm
  gutter: 284    // 0.5 cm 装订线
}
```

---

## 四、LaTeX → Word 公式（OMML）的完整流程

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

### 4.1 第一层：KaTeX 解析 LaTeX → MathML 字符串

```js
const html = katex.renderToString(latex, { output: 'mathml', throwOnError: false });
const div = document.createElement('div');
div.innerHTML = html;
const mathEl = div.querySelector('math');
```

> **关键**：KaTeX 的 MathML 输出结构是 `math > semantics > mrow > [内容]`，
> `semantics` 下还有 `annotation` 节点（含原始 LaTeX 字符串），需要跳过。

### 4.2 第二层：MathML → OMML 递归转换（核心映射表）

| MathML 标签 | docx 对象 | 说明 |
|---|---|---|
| `math`, `mstyle`, `mpadded`, `merror` | 直接展开子节点 | 容器节点 |
| `semantics` | `mathmlNodeToOmml(children[0])` | **必须用 mathmlNodeToOmml 而非 childOmml** |
| `annotation`, `annotation-xml` | 忽略（返回 `[]`） | 含原始 LaTeX，不需要 |
| `mrow` | `tryFencedMrow` 或展开子节点 | 见 §4.3 |
| `mi`, `mn`, `mo`, `mtext` | `new D.MathRun(text)` | 原子文本 |
| `mspace` | `new D.MathRun('\u00a0')` | 空白 |
| `mfrac` | `new D.MathFraction({ numerator, denominator })` | 分数 |
| `msqrt` | `new D.MathRadical({ children })` | 平方根 |
| `mroot` | `new D.MathRadical({ children, degree })` | n次根 |
| `msub` | `new D.MathSubScript({ children, subScript })` | 下标 |
| `msup` | `new D.MathSuperScript({ children, superScript })` | 上标 |
| `msubsup` | `new D.MathSubSuperScript({ children, subScript, superScript })` | 上下标 |
| `mover` (accent=true) | `buildAccent(accentChar, base)` | 见 §4.4 |
| `mover` (accent≠true) | `MathSuperScript` | 极限上标等 |
| `munder` | `MathSubScript` | 极限下标等 |
| `munderover` | `MathSubSuperScript` | 求和上下标 |
| `mtable` | 按行拼接，行间插入 `'; '` | 矩阵/cases 降级处理 |
| 其他 | `childOmml(node)` | 安全展开 |

### 4.3 带定界符的 `\left...\right`（tryFencedMrow）

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

**生成的 OMML**：`<m:d>` 定界符元素，通过 `XmlComponent` 手动构建：
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

### 4.4 重音符号（buildAccent）

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

---

## 五、关键 Bug 记录（已修复）

### Bug 1：所有 `\left...\right` 定界符渲染成三个散字符

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

### Bug 2：重音符号（\widetilde 等）渲染成上标

**现象**：`\widetilde{W}` 在 Word 中显示为 `W~` 而非 $\widetilde{W}$。

**根因**：`mover` 分支未检查 `accent="true"` 属性，统一用 `MathSuperScript` 处理。

**修复**：增加 `accent="true"` 判断，走 `buildAccent` 路径。

---

## 六、内联公式与显示公式的分段解析

段落文字中混合文本和 `$...$` / `$$...$$`，需要先分段再分别处理：

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

---

## 七、样式同步（CSS 变量 ↔ docx 属性）

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

---

## 八、PDF 导出（作为补充）

使用 `html2pdf.js`，直接截图 HTML 预览区：

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

---

## 九、开发调试技巧

1. **在浏览器控制台测试单个 LaTeX**：
   ```js
   // 检查 KaTeX 是否解析正确
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

---

## 十、可直接复用的最简模板

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

## 附录：MathML → OMML 完整覆盖情况

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
> Word 的 OMML `<m:m>` 矩阵与 `<m:eqArr>` 方程数组已完整实现（见第十一节）。

---

## 十一、多行公式自动分行（m:eqArr / m:m）

### 11.1 问题背景

LaTeX 的 `\begin{cases}` / `\begin{aligned}` / `\begin{align}` 等多行环境，
KaTeX 将其渲染为 `<mtable>` 结构（MathML 表格）。  
若直接把各行用分号拼成一行，则 Word 导出后公式失去换行结构，排版混乱。

### 11.2 OMML 多行元素

| OMML 元素 | 用途 | 对应 LaTeX 环境 |
|---|---|---|
| `<m:eqArr>` | 方程数组（每行独立换行） | `cases`, `aligned`, `align`, `gather` |
| `<m:m>` | 矩阵（行+列） | `pmatrix`, `bmatrix`, `vmatrix`, etc. |

**eqArr 结构**：
```xml
<m:eqArr>
  <m:e><!-- 第1行内容 --></m:e>
  <m:e><!-- 第2行内容 --></m:e>
  <m:e><!-- 第3行内容 --></m:e>
</m:eqArr>
```

**matrix 结构**：
```xml
<m:m>
  <m:mr>
    <m:e><!-- cell(0,0) --></m:e>
    <m:e><!-- cell(0,1) --></m:e>
  </m:mr>
  <m:mr>
    <m:e><!-- cell(1,0) --></m:e>
    <m:e><!-- cell(1,1) --></m:e>
  </m:mr>
</m:m>
```

### 11.3 构建函数

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

### 11.4 分发逻辑

**情形1：`<mtable>` 在 fence mrow 内（由 `tryFencedMrow` 处理）**

```js
// tryFencedMrow 中，middleFences.length === 0 分支：
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

**情形2：独立 `<mtable>`（`mathmlNodeToOmml` 的 `mtable` case）**

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

### 11.5 cases 示例（端到端）

LaTeX：`\begin{cases} q^* = \varphi(q^*) \\ Aq^* = 0 \end{cases}`

KaTeX MathML（简化）：
```xml
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

### 11.6 注意事项

- `endChar = ""` 对应 OMML `<m:endChr m:val=""/>` → Word 显示无右括号（cases 正确行为）
- OMML 默认 `m:sepChr` 是 `|`，但单段 `<m:e>` 时 sepChr 无实际渲染影响
- `\\[4pt]` 等行间距修饰：KaTeX 生成 `mspace` 节点，被已有 `mspace` → `\u00a0` 处理；行间距微调在 OMML 中可通过 `<m:eqArrPr>` 的 `<m:lineSp>` 实现（当前版本未实现）
- 嵌套多行（如 `\begin{cases}` 内含 `\begin{aligned}`）：递归调用 `mathmlNodeToOmml` 自然处理

---

## 十二、图片（Image）内容块支持

### 12.1 CONTENT 数据结构

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

### 12.2 HTML 渲染（renderContent）

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

**CSS 样式**：

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

### 12.3 DOCX 导出（exportToDocx）

```js
/** base64 data URL → Uint8Array（docx.js ImageRun 需要 ArrayBuffer/Uint8Array） */
function b64ToUint8(b64src) {
  const b64 = b64src.includes(',') ? b64src.split(',')[1] : b64src;
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

/**
 * Build [imgParagraph, captionParagraph] for a data-URL image.
 * widthPt / heightPt are in pt (docx.js ImageRun uses pt natively via transformation).
 */
function buildImagePara(src, widthPt, heightPt, caption) {
  const imageRun = new D.ImageRun({
    data: b64ToUint8(src),
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
  docChildren.push(...buildImagePara(item.src, wPt, hPt, item.caption || ''));
  break;
}
```

> **docx.js 版本注意**：`ImageRun` 的 `transformation` 属性在 docx.js 7.x 和 8.x 中均支持 pt 单位（内部自动转换为 EMU：1pt = 12700 EMU）。
> 如需 EMU 单位，手动换算：`widthEMU = widthPt * 12700`。

### 12.4 如何从 PDF 提取仿真图片（Python）

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
# 目标宽度440pt，等比例高度
target_w = 440
target_h = round(target_w * h / w)
entry = f"{{type:'img',src:'{src}',caption:'图1. ...',widthPt:{target_w},heightPt:{target_h}}}"
```

### 12.5 从 DOCX 提取图片（Python）

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

