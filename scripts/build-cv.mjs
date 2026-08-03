import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { google } from "googleapis";

const currentFile = fileURLToPath(import.meta.url);
const currentDirectory = path.dirname(currentFile);
const projectRoot = path.resolve(currentDirectory, "..");

const documentId = process.env.GOOGLE_DOC_ID;

const templatePath = path.join(projectRoot, "cv.template.html");
const outputPath = path.join(projectRoot, "cv.html");

if (!documentId) {
  throw new Error(
    "GOOGLE_DOC_ID is missing. Set it before running the CV build."
  );
}

const auth = new google.auth.GoogleAuth({
  scopes: [
    "https://www.googleapis.com/auth/documents.readonly"
  ]
});

const docs = google.docs({
  version: "v1",
  auth
});

const response = await docs.documents.get({
  documentId,
  includeTabsContent: true
});

const document = response.data;

const documentTab = findFirstDocumentTab(
  document.tabs ?? []
);

function getPrimaryHeaderContent(documentTab) {
  const headers = documentTab.headers ?? {};

  const bodyElements =
    documentTab.body?.content ?? [];

  const firstSectionBreak = bodyElements.find(
    (element) => element.sectionBreak
  );

  const sectionStyle =
    firstSectionBreak?.sectionBreak?.sectionStyle ?? {};

  const documentStyle =
    documentTab.documentStyle ?? {};

  const useFirstPageHeader =
    sectionStyle.useFirstPageHeaderFooter ??
    documentStyle.useFirstPageHeaderFooter ??
    false;

  const firstPageHeaderId =
    sectionStyle.firstPageHeaderId ??
    documentStyle.firstPageHeaderId;

  const defaultHeaderId =
    sectionStyle.defaultHeaderId ??
    documentStyle.defaultHeaderId;

  const fallbackHeaderId =
    Object.keys(headers)[0];

  const headerId =
    (
      useFirstPageHeader
        ? firstPageHeaderId
        : defaultHeaderId
    ) ??
    firstPageHeaderId ??
    defaultHeaderId ??
    fallbackHeaderId;

  if (!headerId) {
    return [];
  }

  return headers[headerId]?.content ?? [];
}

function renderCvHeader(elements) {
  const paragraphs = [];

  for (const element of elements) {
    if (!element.paragraph) {
      continue;
    }

    const paragraph = element.paragraph;

    const contents = renderInlineElements(
      paragraph.elements ?? []
    );

    if (!stripHtml(contents).trim()) {
      continue;
    }

    const namedStyle =
      paragraph.paragraphStyle?.namedStyleType ??
      "NORMAL_TEXT";

    paragraphs.push({
      namedStyle,
      contents
    });
  }

  if (paragraphs.length === 0) {
    return "";
  }

  const hasExplicitTitle = paragraphs.some(
    (paragraph) =>
      paragraph.namedStyle === "TITLE"
  );

  let fallbackTitleUsed = false;

  const renderedParagraphs = paragraphs.map(
    ({ namedStyle, contents }) => {
      if (namedStyle === "TITLE") {
        return `<h1 class="cv-name">${contents}</h1>`;
      }

      if (namedStyle === "SUBTITLE") {
        return (
          `<p class="cv-subtitle">` +
          `${contents}</p>`
        );
      }

      /*
       * Fallback for header paragraphs that lost their
       * Google Docs named style.
       */
      if (!hasExplicitTitle && !fallbackTitleUsed) {
        fallbackTitleUsed = true;

        return `<h1 class="cv-name">${contents}</h1>`;
      }

      return (
        `<p class="cv-subtitle">` +
        `${contents}</p>`
      );
    }
  );

  return [
    '<header class="cv-document-header">',
    ...renderedParagraphs,
    "</header>"
  ].join("\n");
}

if (!documentTab) {
  throw new Error(
    "The Google Doc does not contain a readable document tab."
  );
}

const bodyContent =
  documentTab.body?.content ?? [];

const headerContent =
  getPrimaryHeaderContent(documentTab);

const generatedHeaderHtml =
  renderCvHeader(headerContent);

const generatedBodyHtml =
  renderStructuralElements(bodyContent);

const generatedHtml = [
  generatedHeaderHtml,
  generatedBodyHtml
]
  .filter(Boolean)
  .join("\n");

const template = await fs.readFile(templatePath, "utf8");

const marker = "<!-- CV_CONTENT -->";

if (!template.includes(marker)) {
  throw new Error(
    "cv.template.html does not contain <!-- CV_CONTENT -->."
  );
}

const finishedPage = template.replace(
  marker,
  generatedHtml
);

await fs.writeFile(
  outputPath,
  finishedPage,
  "utf8"
);

console.log(
  `Generated cv.html from Google Doc: ${document.title}`
);


/* --------------------------------------------------
   Render document structure
-------------------------------------------------- */
function findFirstDocumentTab(tabs) {
  for (const tab of tabs) {
    if (tab.documentTab) {
      return tab.documentTab;
    }

    const childDocumentTab = findFirstDocumentTab(
      tab.childTabs ?? []
    );

    if (childDocumentTab) {
      return childDocumentTab;
    }
  }

  return null;
}

function renderStructuralElements(elements) {
  const html = [];
  let listOpen = false;

  for (const element of elements) {
    if (element.paragraph) {
      const paragraph = element.paragraph;

     if (paragraph.bullet) {
       if (!listOpen) {
         html.push('<ul class="cv-list">');
         listOpen = true;
       }

       const listItem = renderInlineElements(
         paragraph.elements ?? [],
         {
           trimLeadingWhitespace: true
         }
       );
     
       const nestingLevel = Math.min(
         Math.max(
           Number(
             paragraph.bullet.nestingLevel ?? 0
           ),
           0
         ),
         3
       );

       const listStyleAttribute =
         renderStyleAttribute(
           renderListIndentStyle(paragraph)
         );
     
       if (stripHtml(listItem).trim()) {
         html.push(
           `<li ` +
           `class="cv-list-level-${nestingLevel}"` +
           `${listStyleAttribute}>` +
           `${listItem}</li>`
         );
       }
     
       continue;
     }

      if (listOpen) {
        html.push("</ul>");
        listOpen = false;
      }

      const renderedParagraph = renderParagraph(paragraph);

      if (renderedParagraph) {
        html.push(renderedParagraph);
      }

      continue;
    }

    if (listOpen) {
      html.push("</ul>");
      listOpen = false;
    }

    if (element.table) {
      html.push(renderTable(element.table));
    }

    if (element.horizontalRule) {
      html.push('<hr class="cv-horizontal-rule">');
    }
  }

  if (listOpen) {
    html.push("</ul>");
  }

  return html.join("\n");
}

function dimensionToPoints(dimension) {
  if (!dimension) {
    return null;
  }

  if (dimension.unit !== "PT") {
    return null;
  }

  const magnitude = Number(dimension.magnitude);

  if (!Number.isFinite(magnitude)) {
    return null;
  }

  return magnitude;
}

function formatPoints(value) {
  return Number(value.toFixed(2));
}
function renderParagraphIndentStyle(paragraph) {
  const paragraphStyle =
    paragraph.paragraphStyle ?? {};

  const indentStart =
    dimensionToPoints(
      paragraphStyle.indentStart
    ) ?? 0;

  const indentEnd =
    dimensionToPoints(
      paragraphStyle.indentEnd
    ) ?? 0;

  const firstLinePosition =
    dimensionToPoints(
      paragraphStyle.indentFirstLine
    );

  const textIndent =
    firstLinePosition === null
      ? 0
      : firstLinePosition - indentStart;

  return [
    `--cv-indent-start: ${formatPoints(indentStart)}pt`,
    `--cv-indent-end: ${formatPoints(indentEnd)}pt`,
    `--cv-text-indent: ${formatPoints(textIndent)}pt`
  ].join("; ");
}

function renderParagraph(paragraph) {
  const contents = renderInlineElements(
    paragraph.elements ?? [],
    {
      trimLeadingWhitespace: true
    }
  );

  if (!stripHtml(contents).trim()) {
    return "";
  }

  const namedStyle =
    paragraph.paragraphStyle?.namedStyleType ??
    "NORMAL_TEXT";

  const styleAttribute =
    renderStyleAttribute(
      renderParagraphIndentStyle(paragraph)
    );

  switch (namedStyle) {
    case "TITLE":
      return (
        `<h1 class="cv-name"` +
        `${styleAttribute}>` +
        `${contents}</h1>`
      );

    case "SUBTITLE":
      return (
        `<p class="cv-subtitle"` +
        `${styleAttribute}>` +
        `${contents}</p>`
      );

    case "HEADING_3":
      return (
        `<h2 class="cv-section-title"` +
        `${styleAttribute}>` +
        `${contents}</h2>`
      );

    case "HEADING_4":
      return (
        `<h3 class="cv-entry-title"` +
        `${styleAttribute}>` +
        `${contents}</h3>`
      );

    default:
      return (
        `<p class="cv-paragraph"` +
        `${styleAttribute}>` +
        `${contents}</p>`
      );
  }
}

function renderListIndentStyle(paragraph) {
  const indentStart =
    dimensionToPoints(
      paragraph.paragraphStyle?.indentStart
    );

  if (indentStart === null) {
    return "";
  }

  return (
    `--cv-list-indent: ` +
    `${formatPoints(indentStart)}pt`
  );
}

function renderStyleAttribute(styleValue) {
  if (!styleValue) {
    return "";
  }

  return (
    ` style="${escapeAttribute(styleValue)}"`
  );
}


/* --------------------------------------------------
   Render inline formatting
-------------------------------------------------- */

function renderInlineElements(
  elements,
  options = {}
) {
  const {
    trimLeadingWhitespace = false
  } = options;

  const finalTextRun =
    findFinalTextRun(elements);

  let paragraphHasStarted = false;

  return elements
    .map((element, index) => {
      const textRun = element.textRun;

      if (!textRun?.content) {
        return "";
      }

      let content = textRun.content;

      /*
       * Google Docs ends paragraphs with a newline.
       * Remove only the final paragraph newline.
       */
      if (index === finalTextRun) {
        content = content.replace(/\n$/, "");
      }

      /*
       * Remove tabs and spaces used merely to simulate
       * paragraph indentation. The actual indentation
       * now comes from paragraphStyle.indentStart.
       */
      if (
        trimLeadingWhitespace &&
        !paragraphHasStarted
      ) {
        content = content.replace(
          /^[\t ]+/,
          ""
        );
      }

      if (!content) {
        return "";
      }

      if (
        content.replace(/\n/g, "").length > 0
      ) {
        paragraphHasStarted = true;
      }

      let rendered = escapeHtml(content)
        .replaceAll("\t", "    ")
        .replaceAll("\n", "<br>");

      const style = textRun.textStyle ?? {};

      if (style.bold) {
        rendered =
          `<strong>${rendered}</strong>`;
      }

      if (style.italic) {
        rendered =
          `<em>${rendered}</em>`;
      }

      const link =
        validateLink(style.link?.url);

      if (link) {
        rendered =
          `<a href="${escapeAttribute(link)}">` +
          `${rendered}</a>`;
      }

      return rendered;
    })
    .join("");
}

function findFinalTextRun(elements) {
  for (
    let index = elements.length - 1;
    index >= 0;
    index -= 1
  ) {
    if (elements[index].textRun?.content) {
      return index;
    }
  }

  return -1;
}


/* --------------------------------------------------
   Render tables
-------------------------------------------------- */

function renderTable(table) {
  const rows = table.tableRows ?? [];

  const renderedRows = rows.map((row) => {
    const cells = row.tableCells ?? [];

    const renderedCells = cells.map((cell) => {
      const contents = renderStructuralElements(
        cell.content ?? []
      );

      return `<td>${contents}</td>`;
    });

    return `<tr>${renderedCells.join("")}</tr>`;
  });

  return (
    '<table class="cv-table">' +
    "<tbody>" +
    renderedRows.join("") +
    "</tbody>" +
    "</table>"
  );
}


/* --------------------------------------------------
   Safety and escaping
-------------------------------------------------- */

function validateLink(url) {
  if (!url) {
    return "";
  }

  try {
    const parsedUrl = new URL(url);

    const permittedProtocols = [
      "http:",
      "https:",
      "mailto:"
    ];

    if (permittedProtocols.includes(parsedUrl.protocol)) {
      return parsedUrl.href;
    }
  } catch {
    return "";
  }

  return "";
}

function stripHtml(value) {
  return value.replace(/<[^>]*>/g, "");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value)
    .replaceAll("`", "&#096;");
}