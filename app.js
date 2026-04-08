const sourceInput = document.getElementById("sourceInput");
const preview = document.getElementById("preview");
const statusEl = document.getElementById("status");
const parseBtn = document.getElementById("parseBtn");
const copyWordBtn = document.getElementById("copyWordBtn");
const exportDocxBtn = document.getElementById("exportDocxBtn");
const copyPlainBtn = document.getElementById("copyPlainBtn");
const downloadWordBtn = document.getElementById("downloadWordBtn");
const clearBtn = document.getElementById("clearBtn");
const MML2OMML_ESM_URL = "https://cdn.jsdelivr.net/npm/mathml2omml/+esm";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const BUILD_ID = "2026-04-08-docx-fix3";
let mml2ommlPromise = null;

marked.setOptions({
  gfm: true,
  breaks: true
});

parseBtn.addEventListener("click", parseAndRender);
copyWordBtn.addEventListener("click", copyForWord);
exportDocxBtn.addEventListener("click", exportDocx);
copyPlainBtn.addEventListener("click", copyPlainText);
downloadWordBtn.addEventListener("click", downloadWordReadyFile);
clearBtn.addEventListener("click", clearAll);
setStatus(`Sẵn sàng (${BUILD_ID}). Dán nội dung vào khung bên trái rồi bấm Parse + Preview.`);

sourceInput.addEventListener("keydown", (event) => {
  const isParseShortcut = (event.ctrlKey || event.metaKey) && event.key === "Enter";
  if (isParseShortcut) {
    event.preventDefault();
    parseAndRender();
  }
});

function parseAndRender() {
  const normalized = normalizeInput(sourceInput.value);
  if (!normalized.trim()) {
    preview.innerHTML = "";
    setStatus("Chưa có nội dung để parse.");
    return;
  }

  const extracted = extractMathSegments(normalized);
  const markdownReady = extracted.textWithPlaceholders.replace(/\\([{}])/g, "$1");
  const markdownHtml = marked.parse(markdownReady);
  const htmlWithMath = injectMath(extracted.tokens, markdownHtml);
  const safeHtml = sanitizeWithMathMl(htmlWithMath);

  preview.innerHTML = safeHtml;
  setStatus(
    `Đã parse xong: ${extracted.tokens.length} công thức. Bạn có thể chỉnh tay bên phải rồi copy sang Word.`
  );
}

function normalizeInput(raw) {
  return raw
    .replace(/\r\n?/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/[‐‑‒–—]/g, "-");
}

function extractMathSegments(text) {
  const tokens = [];
  let textWithPlaceholders = "";
  let i = 0;

  while (i < text.length) {
    const blockSlash = tryDelimitedBySlash(text, i, "\\[", "\\]", true);
    if (blockSlash) {
      const placeholder = newPlaceholder(tokens.length);
      tokens.push({
        ...blockSlash,
        placeholder
      });
      textWithPlaceholders += placeholder;
      i += blockSlash.raw.length;
      continue;
    }

    const inlineSlash = tryDelimitedBySlash(text, i, "\\(", "\\)", false);
    if (inlineSlash) {
      const placeholder = newPlaceholder(tokens.length);
      tokens.push({
        ...inlineSlash,
        placeholder
      });
      textWithPlaceholders += placeholder;
      i += inlineSlash.raw.length;
      continue;
    }

    const blockDollar = tryDelimitedByDollar(text, i, "$$", true);
    if (blockDollar) {
      const placeholder = newPlaceholder(tokens.length);
      tokens.push({
        ...blockDollar,
        placeholder
      });
      textWithPlaceholders += placeholder;
      i += blockDollar.raw.length;
      continue;
    }

    const inlineDollar = tryDelimitedByDollar(text, i, "$", false);
    if (inlineDollar) {
      const placeholder = newPlaceholder(tokens.length);
      tokens.push({
        ...inlineDollar,
        placeholder
      });
      textWithPlaceholders += placeholder;
      i += inlineDollar.raw.length;
      continue;
    }

    textWithPlaceholders += text[i];
    i += 1;
  }

  return {
    tokens,
    textWithPlaceholders
  };
}

function tryDelimitedBySlash(text, startIndex, startDelimiter, endDelimiter, display) {
  if (!text.startsWith(startDelimiter, startIndex)) {
    return null;
  }

  const contentStart = startIndex + startDelimiter.length;
  const endIndex = text.indexOf(endDelimiter, contentStart);
  if (endIndex === -1) {
    return null;
  }

  const tex = text.slice(contentStart, endIndex);
  if (!tex.trim()) {
    return null;
  }

  const raw = text.slice(startIndex, endIndex + endDelimiter.length);
  return {
    tex,
    display,
    raw
  };
}

function tryDelimitedByDollar(text, startIndex, startDelimiter, display) {
  if (!text.startsWith(startDelimiter, startIndex)) {
    return null;
  }

  if (startDelimiter === "$" && text.startsWith("$$", startIndex)) {
    return null;
  }

  if (isEscaped(text, startIndex)) {
    return null;
  }

  if (startDelimiter === "$$") {
    const contentStart = startIndex + 2;
    for (let i = contentStart; i < text.length - 1; i += 1) {
      const isClose = text[i] === "$" && text[i + 1] === "$" && !isEscaped(text, i);
      if (isClose) {
        const tex = text.slice(contentStart, i);
        if (!tex.trim()) {
          return null;
        }
        const raw = text.slice(startIndex, i + 2);
        return {
          tex,
          display,
          raw
        };
      }
    }
    return null;
  }

  for (let i = startIndex + 1; i < text.length; i += 1) {
    if (text[i] === "\n") {
      break;
    }
    if (text[i] === "$" && !isEscaped(text, i)) {
      const tex = text.slice(startIndex + 1, i);
      if (!tex.trim()) {
        return null;
      }
      const raw = text.slice(startIndex, i + 1);
      return {
        tex,
        display,
        raw
      };
    }
  }

  return null;
}

function isEscaped(text, index) {
  let slashCount = 0;
  for (let i = index - 1; i >= 0 && text[i] === "\\"; i -= 1) {
    slashCount += 1;
  }
  return slashCount % 2 === 1;
}

function injectMath(tokens, markdownHtml) {
  let withMath = markdownHtml;
  for (const token of tokens) {
    withMath = withMath.split(token.placeholder).join(renderMathToken(token));
  }
  return withMath;
}

function renderMathToken(token) {
  if (!window.temml || typeof window.temml.renderToString !== "function") {
    return asMathError(token.raw, "Thư viện temml chưa sẵn sàng.");
  }

  try {
    return window.temml.renderToString(token.tex, {
      displayMode: token.display,
      throwOnError: true,
      annotate: false
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Không parse được công thức.";
    return asMathError(token.raw, message);
  }
}

function asMathError(raw, message) {
  return `<span class="math-error" title="${escapeHtml(message)}">${escapeHtml(raw)}</span>`;
}

function sanitizeWithMathMl(html) {
  return window.DOMPurify.sanitize(html, {
    USE_PROFILES: {
      html: true,
      mathMl: true
    }
  });
}

async function copyForWord() {
  if (!preview.innerHTML.trim()) {
    setStatus("Chưa có nội dung đã parse để copy.");
    return;
  }

  const { fragment: wordFragment, usedOmml } = await buildWordFragment(preview.innerHTML);
  const wordHtml = buildWordHtml(wordFragment);
  const plainText = buildPlainTextForWord();

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const item = new ClipboardItem({
        "text/html": new Blob([wordHtml], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      });
      await navigator.clipboard.write([item]);
      setStatus(
        usedOmml
          ? "Đã copy bản OMML cho Word. Bạn vào Word và dán bằng Ctrl + V."
          : "Đã copy nhưng chưa bật được OMML (đang dùng fallback). Hãy dùng nút Xuất DOCX chuẩn."
      );
      return;
    }

    const legacyCopied = legacyCopyHtml(wordFragment);
    if (legacyCopied) {
      setStatus("Đã copy bằng chế độ tương thích. Nếu công thức chưa chuẩn, hãy dùng nút tải file.");
      return;
    }

    await navigator.clipboard.writeText(plainText);
    setStatus("Trình duyệt không hỗ trợ copy HTML đầy đủ, đã copy dạng text thuần.");
  } catch (error) {
    const fallback = legacyCopyHtml(wordFragment);
    if (fallback) {
      setStatus("Đã copy bằng chế độ dự phòng.");
      return;
    }

    const message = error instanceof Error ? error.message : "Lỗi clipboard.";
    setStatus(`Không copy được tự động (${message}). Hãy dùng nút tải file.`);
  }
}

async function copyPlainText() {
  const plain = buildPlainTextForWord();
  if (!plain.trim()) {
    setStatus("Không có text để copy.");
    return;
  }

  try {
    await navigator.clipboard.writeText(plain);
    setStatus("Đã copy text thuần.");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Không xác định.";
    setStatus(`Copy text thuần thất bại: ${message}`);
  }
}

async function downloadWordReadyFile() {
  if (!preview.innerHTML.trim()) {
    setStatus("Chưa có nội dung để tải.");
    return;
  }

  const { fragment: wordFragment, usedOmml } = await buildWordFragment(preview.innerHTML);
  const wordHtml = buildWordHtml(wordFragment);
  const blob = new Blob([wordHtml], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "chatgpt-to-word.doc";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  setStatus(
    usedOmml
      ? "Đã tải file .doc có OMML. Mở bằng Word để giữ công thức chuẩn."
      : "Đã tải file .doc fallback. Nếu công thức chưa chuẩn, hãy kiểm tra mạng rồi thử lại."
  );
}

async function exportDocx() {
  const normalized = normalizeInput(sourceInput.value);
  if (!normalized.trim()) {
    setStatus("Chưa có nội dung để xuất DOCX.");
    return;
  }

  if (!window.JSZip) {
    setStatus("Thiếu thư viện JSZip để tạo DOCX. Hãy tải lại trang bằng Ctrl + F5.");
    return;
  }

  if (!window.temml || typeof window.temml.renderToString !== "function") {
    setStatus("Thiếu thư viện temml. Hãy đợi tải xong trang rồi thử lại.");
    return;
  }

  const mml2omml = await getMml2OmmlConverter();
  if (!mml2omml) {
    setStatus("Không tải được bộ chuyển MathML -> OMML. Kiểm tra mạng rồi thử lại.");
    return;
  }

  try {
    const documentXml = buildDocxDocumentXml(normalized, mml2omml);
    if (!isWellFormedXml(documentXml)) {
      setStatus("Không tạo được DOCX hợp lệ từ nội dung hiện tại. Hãy thử rút gọn công thức và xuất lại.");
      return;
    }
    const zip = new window.JSZip();

    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`
    );

    zip.folder("_rels").file(
      ".rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
    );

    const nowIso = new Date().toISOString();
    zip.folder("docProps").file(
      "core.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ChatGPT Formula Export</dc:title>
  <dc:creator>TransformLatex</dc:creator>
  <cp:lastModifiedBy>TransformLatex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:modified>
</cp:coreProperties>`
    );

    zip.folder("docProps").file(
      "app.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>TransformLatex</Application>
</Properties>`
    );

    const wordFolder = zip.folder("word");
    wordFolder.file("document.xml", documentXml);
    wordFolder.file(
      "styles.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr/>
    <w:rPr/>
  </w:style>
</w:styles>`
    );
    wordFolder.file(
      "settings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:mathPr/>
</w:settings>`
    );
    wordFolder.folder("_rels").file(
      "document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`
    );

    const blob = await zip.generateAsync({ type: "blob", mimeType: DOCX_MIME });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "chatgpt-formula.docx";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    setStatus("Đã xuất DOCX chuẩn công thức. Mở file này bằng Word sẽ hiển thị phân số đúng.");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Không xác định";
    setStatus(`Xuất DOCX thất bại: ${message}`);
  }
}

function buildWordHtml(fragment) {
  const styles = `
<style>
  body {
    font-family: "Times New Roman", serif;
    font-size: 12pt;
    line-height: 1.45;
    color: #111;
  }
  p { margin: 0 0 8pt; }
  ul, ol { margin: 0 0 8pt 22pt; padding: 0; }
  li { margin: 0 0 4pt; }
  math {
    font-family: "Cambria Math", "Times New Roman", serif;
  }
  math[display="block"] {
    margin: 8pt 0;
  }
</style>`;

  return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8">
  <meta name="ProgId" content="Word.Document">
  ${styles}
</head>
<body>
  ${fragment}
</body>
</html>`;
}

async function buildWordFragment(fragment) {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = fragment;

  wrapper.querySelectorAll(".math-error").forEach((node) => {
    const text = node.textContent || "";
    node.replaceWith(document.createTextNode(text));
  });

  wrapper.querySelectorAll("annotation, annotation-xml").forEach((node) => {
    node.remove();
  });

  wrapper.querySelectorAll("semantics").forEach((semanticsNode) => {
    const firstMathNode = Array.from(semanticsNode.children).find((child) => {
      const tag = child.tagName.toLowerCase();
      return tag !== "annotation" && tag !== "annotation-xml";
    });

    if (firstMathNode) {
      semanticsNode.replaceWith(firstMathNode);
      return;
    }

    semanticsNode.remove();
  });

  wrapper.querySelectorAll("math").forEach((mathNode) => {
    if (!mathNode.getAttribute("xmlns")) {
      mathNode.setAttribute("xmlns", "http://www.w3.org/1998/Math/MathML");
    }
  });

  const mathNodes = Array.from(wrapper.querySelectorAll("math"));
  const mathPlaceholders = [];
  for (let i = 0; i < mathNodes.length; i += 1) {
    const mathNode = mathNodes[i];
    const placeholder = `@@OMML_TOKEN_${i}_${Math.random().toString(36).slice(2, 8)}@@`;
    const isDisplay = mathNode.getAttribute("display") === "block";
    const mathMl = new XMLSerializer().serializeToString(mathNode);

    mathNode.replaceWith(document.createTextNode(placeholder));
    mathPlaceholders.push({
      placeholder,
      isDisplay,
      mathMl
    });
  }

  const walker = document.createTreeWalker(wrapper, NodeFilter.SHOW_TEXT);
  const textNodes = [];
  while (walker.nextNode()) {
    textNodes.push(walker.currentNode);
  }

  for (const textNode of textNodes) {
    const parent = textNode.parentElement;
    if (parent && parent.closest("math")) {
      continue;
    }
    textNode.textContent = (textNode.textContent || "").replace(/\\([{}])/g, "$1");
  }

  let html = wrapper.innerHTML;
  let usedOmml = false;
  const mml2omml = await getMml2OmmlConverter();
  if (!mml2omml) {
    for (const entry of mathPlaceholders) {
      html = html.replace(entry.placeholder, entry.mathMl);
    }
    return {
      fragment: html,
      usedOmml
    };
  }

  for (const entry of mathPlaceholders) {
    const omml = convertMathMlToOmml(mml2omml, entry.mathMl, entry.isDisplay);
    if (omml) {
      html = html.replace(entry.placeholder, omml);
      usedOmml = true;
    } else {
      html = html.replace(entry.placeholder, entry.mathMl);
    }
  }

  return {
    fragment: html,
    usedOmml
  };
}

function buildDocxDocumentXml(normalizedInput, mml2omml) {
  const extracted = extractMathSegments(normalizedInput);
  const plainReady = extracted.textWithPlaceholders.replace(/\\([{}])/g, "$1");
  const lines = plainReady.split("\n");
  const tokensByPlaceholder = new Map(
    extracted.tokens.map((token) => [token.placeholder, token])
  );
  const ommlCache = new Map();
  const paragraphs = lines
    .map((line) => buildDocxParagraphXml(line, tokensByPlaceholder, mml2omml, ommlCache))
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    ${paragraphs || "<w:p/>"}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function buildDocxParagraphXml(line, tokensByPlaceholder, mml2omml, ommlCache) {
  if (!line.trim()) {
    return "<w:p/>";
  }

  const pieces = splitLineByMathPlaceholders(line, tokensByPlaceholder);
  const onlyDisplayMath =
    pieces.length === 1 && pieces[0].type === "math" && pieces[0].token.display;
  if (onlyDisplayMath) {
    const omml = convertTokenToInlineOmml(pieces[0].token, mml2omml, ommlCache);
    if (omml) {
      return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${omml}</w:p>`;
    }
    return `<w:p>${textRunXml(pieces[0].token.raw)}</w:p>`;
  }

  const content = [];
  for (const piece of pieces) {
    if (piece.type === "text") {
      const run = textRunXml(piece.value);
      if (run) {
        content.push(run);
      }
      continue;
    }

    const omml = convertTokenToInlineOmml(piece.token, mml2omml, ommlCache);
    if (omml) {
      content.push(omml);
    } else {
      content.push(textRunXml(piece.token.raw));
    }
  }

  if (content.length === 0) {
    return "<w:p/>";
  }

  return `<w:p>${content.join("")}</w:p>`;
}

function splitLineByMathPlaceholders(line, tokensByPlaceholder) {
  const parts = [];
  const regex = /@@MATH_TOKEN_[0-9]+_[a-z0-9]+@@/g;
  let lastIndex = 0;
  let match = null;

  while ((match = regex.exec(line)) !== null) {
    const tokenText = match[0];
    if (match.index > lastIndex) {
      parts.push({
        type: "text",
        value: line.slice(lastIndex, match.index)
      });
    }

    const token = tokensByPlaceholder.get(tokenText);
    if (token) {
      parts.push({
        type: "math",
        token
      });
    } else {
      parts.push({
        type: "text",
        value: tokenText
      });
    }

    lastIndex = match.index + tokenText.length;
  }

  if (lastIndex < line.length) {
    parts.push({
      type: "text",
      value: line.slice(lastIndex)
    });
  }

  if (parts.length === 0) {
    parts.push({
      type: "text",
      value: line
    });
  }

  return parts;
}

function convertTokenToInlineOmml(token, mml2omml, ommlCache) {
  if (ommlCache.has(token.placeholder)) {
    return ommlCache.get(token.placeholder);
  }

  let inlineOmml = "";
  try {
    const mathMlRaw = window.temml.renderToString(token.tex, {
      displayMode: token.display,
      throwOnError: true,
      annotate: false
    });
    const mathMl = ensureMathMlNamespace(mathMlRaw);
    const ommlRaw = convertMathMlToOmml(mml2omml, mathMl, false);
    inlineOmml = validateAndNormalizeOmml(normalizeOmmlInline(ommlRaw));
    if (!inlineOmml) {
      inlineOmml = buildLinearOmml(token.tex);
    }
  } catch {
    inlineOmml = buildLinearOmml(token.tex);
  }

  ommlCache.set(token.placeholder, inlineOmml);
  return inlineOmml;
}

function normalizeOmmlInline(omml) {
  if (!omml || typeof omml !== "string") {
    return "";
  }

  const trimmed = omml.trim();
  if (trimmed.startsWith("<m:oMathPara")) {
    const inner = trimmed.match(/<m:oMath\b[\s\S]*<\/m:oMath>/);
    return inner ? inner[0] : "";
  }

  if (trimmed.startsWith("<m:oMath")) {
    return trimmed;
  }

  return "";
}

function validateAndNormalizeOmml(omml) {
  if (!omml) {
    return "";
  }

  const parser = new DOMParser();
  const wrapped = `<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${omml}</root>`;
  const parsed = parser.parseFromString(wrapped, "application/xml");
  if (parsed.getElementsByTagName("parsererror").length > 0) {
    return "";
  }

  const first = parsed.documentElement.firstElementChild;
  if (!first || first.tagName.toLowerCase() !== "m:omath") {
    return "";
  }

  return new XMLSerializer().serializeToString(first);
}

function buildLinearOmml(tex) {
  const linear = texToLinearPlain(tex) || tex;
  return `<m:oMath><m:r><m:t xml:space="preserve">${escapeXml(linear)}</m:t></m:r></m:oMath>`;
}

function ensureMathMlNamespace(mathMl) {
  if (!mathMl || typeof mathMl !== "string") {
    return "";
  }

  if (/^<math\b/.test(mathMl) && !/^<math\b[^>]*\bxmlns=/.test(mathMl)) {
    return mathMl.replace(
      /^<math\b/,
      '<math xmlns="http://www.w3.org/1998/Math/MathML"'
    );
  }

  return mathMl;
}

function textRunXml(text) {
  if (!text) {
    return "";
  }

  const normalizedText = text.replace(/\t/g, "    ");
  return `<w:r><w:t xml:space="preserve">${escapeXml(normalizedText)}</w:t></w:r>`;
}

function buildPlainTextForWord() {
  const normalized = normalizeInput(sourceInput.value);
  if (!normalized.trim()) {
    return preview.innerText || "";
  }

  const extracted = extractMathSegments(normalized);
  let plain = extracted.textWithPlaceholders.replace(/\\([{}])/g, "$1");

  for (const token of extracted.tokens) {
    const linear = texToLinearPlain(token.tex);
    plain = plain.split(token.placeholder).join(linear);
  }

  return plain;
}

function texToLinearPlain(tex) {
  let plain = tex;
  let previous = "";

  // Unroll common \frac patterns repeatedly to handle simple nesting.
  while (plain !== previous) {
    previous = plain;
    plain = plain.replace(/\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}/g, "($1)/($2)");
  }

  const replacements = [
    [/\\cdot/g, "·"],
    [/\\times/g, "×"],
    [/\\leq/g, "<="],
    [/\\geq/g, ">="],
    [/\\neq/g, "!="],
    [/\\infty/g, "∞"],
    [/\\alpha/g, "α"],
    [/\\beta/g, "β"],
    [/\\gamma/g, "γ"],
    [/\\delta/g, "δ"],
    [/\\theta/g, "θ"],
    [/\\lambda/g, "λ"],
    [/\\mu/g, "μ"],
    [/\\pi/g, "π"],
    [/\\sigma/g, "σ"],
    [/\\omega/g, "ω"],
    [/\\sum/g, "Σ"],
    [/\\int/g, "∫"],
    [/\\to/g, "→"],
    [/\\pm/g, "±"],
    [/\\sqrt\s*\{([^{}]+)\}/g, "sqrt($1)"],
    [/\\left/g, ""],
    [/\\right/g, ""],
    [/\\,/g, " "],
    [/\\;/g, " "],
    [/\\:/g, " "],
    [/\\!/g, ""]
  ];

  for (const [pattern, value] of replacements) {
    plain = plain.replace(pattern, value);
  }

  plain = plain.replace(/[{}]/g, "");
  plain = plain.replace(/\\[a-zA-Z]+/g, "");
  plain = plain.replace(/\s+/g, " ").trim();
  return plain;
}

async function getMml2OmmlConverter() {
  if (!mml2ommlPromise) {
    mml2ommlPromise = import(MML2OMML_ESM_URL)
      .then((mod) => (typeof mod.mml2omml === "function" ? mod.mml2omml : null))
      .catch(() => {
        mml2ommlPromise = null;
        return null;
      });
  }
  return mml2ommlPromise;
}

function convertMathMlToOmml(mml2omml, mathMl, isDisplay) {
  try {
    const ommlRaw = mml2omml(mathMl);
    if (!ommlRaw || typeof ommlRaw !== "string") {
      return "";
    }

    const omml = ommlRaw.trim();
    if (!omml.startsWith("<m:")) {
      return "";
    }

    if (isDisplay) {
      if (omml.startsWith("<m:oMathPara")) {
        return omml;
      }
      return `<m:oMathPara>${omml}</m:oMathPara>`;
    }

    return omml;
  } catch {
    return "";
  }
}

function legacyCopyHtml(fragment) {
  const helper = document.createElement("div");
  helper.setAttribute("contenteditable", "true");
  helper.style.position = "fixed";
  helper.style.left = "-9999px";
  helper.style.top = "0";
  helper.innerHTML = fragment;
  document.body.appendChild(helper);

  const selection = window.getSelection();
  if (!selection) {
    helper.remove();
    return false;
  }

  const range = document.createRange();
  range.selectNodeContents(helper);
  selection.removeAllRanges();
  selection.addRange(range);
  const copied = document.execCommand("copy");
  selection.removeAllRanges();
  helper.remove();
  return copied;
}

function clearAll() {
  sourceInput.value = "";
  preview.innerHTML = "";
  setStatus("Đã xóa toàn bộ nội dung.");
  sourceInput.focus();
}

function setStatus(message) {
  statusEl.textContent = message;
}

function newPlaceholder(index) {
  return `@@MATH_TOKEN_${index}_${Math.random().toString(36).slice(2, 8)}@@`;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeXml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function isWellFormedXml(xmlText) {
  if (!xmlText || typeof xmlText !== "string") {
    return false;
  }

  const parser = new DOMParser();
  const parsed = parser.parseFromString(xmlText, "application/xml");
  return parsed.getElementsByTagName("parsererror").length === 0;
}
