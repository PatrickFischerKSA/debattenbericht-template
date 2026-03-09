const form = document.getElementById("reportForm");
const exportButton = document.getElementById("exportButton");
const sampleButton = document.getElementById("sampleButton");
const readinessList = document.getElementById("readinessList");

const preview = {
  authorsInline: document.getElementById("previewAuthorsInline"),
  dateInline: document.getElementById("previewDateInline"),
  dateBadge: document.getElementById("previewDateBadge"),
  theme: document.getElementById("previewTheme"),
  title: document.getElementById("previewTitle"),
  subtitle: document.getElementById("previewSubtitle"),
  lead: document.getElementById("previewLead"),
  debaters: document.getElementById("previewDebaters"),
  quote: document.getElementById("previewQuote"),
  quoteSpeaker: document.getElementById("previewQuoteSpeaker"),
  blockOne: document.getElementById("previewBlockOne"),
  blockTwo: document.getElementById("previewBlockTwo"),
  authorCaption: document.getElementById("previewAuthorCaption"),
  debateCaption: document.getElementById("previewDebateCaption"),
  authorImage: document.getElementById("previewAuthorImage"),
  debateImage: document.getElementById("previewDebateImage"),
  authorPlaceholder: document.getElementById("authorImagePlaceholder"),
  debatePlaceholder: document.getElementById("debateImagePlaceholder")
};

const defaultState = {
  reportDate: toInputDate(new Date()),
  debateTheme: "",
  debaterNames: "",
  authorNames: "",
  title: "",
  subtitle: "",
  lead: "",
  blockOne: "",
  blockTwo: "",
  debateCaption: "",
  quoteText: "",
  quoteSpeaker: "",
  authorImageDataUrl: "",
  debateImageDataUrl: ""
};

const state = { ...defaultState };

const requiredFields = [
  "authorNames",
  "title",
  "subtitle",
  "lead",
  "blockOne",
  "blockTwo",
  "quoteText",
  "quoteSpeaker"
];

function toInputDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDisplayDate(raw) {
  if (!raw) return "Datum";
  const date = new Date(`${raw}T12:00:00`);
  if (Number.isNaN(date.getTime())) return raw;
  return new Intl.DateTimeFormat("de-CH", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).format(date);
}

function splitSentences(text) {
  const cleaned = text.trim();
  if (!cleaned) return [];
  if (typeof Intl !== "undefined" && Intl.Segmenter) {
    const segmenter = new Intl.Segmenter("de", { granularity: "sentence" });
    return Array.from(segmenter.segment(cleaned), (entry) => entry.segment.trim()).filter(Boolean);
  }
  return cleaned
    .split(/(?<=[.!?])\s+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function splitWords(text) {
  return (text.match(/[A-Za-zÄÖÜäöü0-9-]+/g) || []).filter(Boolean);
}

function countChars(text) {
  return [...(text || "")].length;
}

function escapeHtml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeXml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function paragraphsToHtml(text, placeholder) {
  const chunks = String(text || "")
    .split(/\n{2,}/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (!chunks.length) {
    return `<p>${escapeHtml(placeholder)}</p>`;
  }

  return chunks
    .map((part) => `<p>${escapeHtml(part).replace(/\n/g, "<br>")}</p>`)
    .join("");
}

function getFirstNames() {
  return state.debaterNames
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => entry.split(/\s+/)[0]);
}

function getThemeKeywords() {
  return state.debateTheme
    .split(/\s+/)
    .map((entry) => entry.trim().toLowerCase())
    .filter((entry) => entry.length >= 4);
}

function hasRepeatedWord(text) {
  return /\b([A-Za-zÄÖÜäöü-]{3,})\s+\1\b/i.test(text);
}

function analyzeStyle(text) {
  const sentences = splitSentences(text);
  const longSentences = [];
  const passiveSentences = [];

  sentences.forEach((sentence) => {
    const wordCount = splitWords(sentence).length;
    if (wordCount > 24) {
      longSentences.push(`${wordCount} Wörter: ${sentence.slice(0, 90)}${sentence.length > 90 ? "..." : ""}`);
    }

    if (/(wird|werden|wurde|wurden|worden|ist|sind|war|waren)\s+[A-Za-zÄÖÜäöü-]+(?:t|en|iert)\b/i.test(sentence)) {
      passiveSentences.push(sentence.slice(0, 90) + (sentence.length > 90 ? "..." : ""));
    }
  });

  return {
    sentenceCount: sentences.length,
    longSentences,
    passiveSentences,
    repeatedWord: hasRepeatedWord(text),
    lowercaseStart: /(?:^|[.!?]\s+)[a-zäöü]/.test(text.trim())
  };
}

function analyzeLead(text) {
  const lower = text.toLowerCase();
  const names = getFirstNames();
  const coverage = {
    wer: names.some((name) => lower.includes(name.toLowerCase())) || /(schüler|schueler|publikum|team|debattier)/i.test(text),
    was:
      /(debatte|diskut|frage|thema|abstimmung|kontroverse)/i.test(text) ||
      getThemeKeywords().some((keyword) => lower.includes(keyword)),
    wann: /(heute|gestern|morgen|am\s+\w+tag|montag|dienstag|mittwoch|donnerstag|freitag|samstag|sonntag|\d{1,2}\.\d{1,2}\.\d{2,4})/i.test(text),
    wo: /(in|im|an|auf|bei)\s+[A-ZÄÖÜ][A-Za-zÄÖÜäöü-]+/.test(text),
    warum: /(weil|deshalb|darum|wegen|zur frage|um\s+zu|begründ)/i.test(text),
    wie: /(hitzig|sachlich|knapp|deutlich|eng|ruhig|lebhaft|fazit|bilanz|am ende|schlussendlich|überzeugend|ueberzeugend)/i.test(text)
  };
  return coverage;
}

function feedbackLine(level, message) {
  return `<div class="feedback-line ${level}">${message}</div>`;
}

function setFeedback(field, lines) {
  const target = document.querySelector(`[data-feedback-for="${field}"]`);
  if (!target) return;
  target.innerHTML = lines.join("");
}

function validateTextField(value, options = {}) {
  const lines = [];
  const length = countChars(value);
  const max = options.maxChars;
  const min = options.minChars || 0;

  if (!value.trim()) {
    lines.push(feedbackLine(options.required ? "error" : "warn", "Eingabe fehlt."));
    return { ok: !options.required, lines };
  }

  if (max) {
    const remaining = max - length;
    if (remaining < 0) {
      lines.push(feedbackLine("error", `${length}/${max} Zeichen. Limit überschritten.`));
    } else if (remaining <= Math.max(8, Math.round(max * 0.08))) {
      lines.push(feedbackLine("warn", `${length}/${max} Zeichen. Knapp am Limit.`));
    } else {
      lines.push(feedbackLine("ok", `${length}/${max} Zeichen. Formales Limit eingehalten.`));
    }
  }

  if (min && length < min) {
    lines.push(feedbackLine("warn", `Noch knapp: ${length} Zeichen von empfohlenen ${min}.`));
  }

  const style = analyzeStyle(value);
  if (style.longSentences.length) {
    lines.push(
      feedbackLine(
        "warn",
        `${style.longSentences.length} lange Sätze erkannt. Ziel: möglichst unter 25 Wörtern pro Satz.`
      )
    );
  } else if (style.sentenceCount > 0) {
    lines.push(feedbackLine("ok", "Satzlängen wirken kompakt."));
  }

  if (style.passiveSentences.length) {
    lines.push(feedbackLine("warn", `${style.passiveSentences.length} Passiv-Muster erkannt. Aktiv formulieren.`));
  } else if (style.sentenceCount > 0) {
    lines.push(feedbackLine("ok", "Aktivstil wirkt stimmig."));
  }

  if (style.repeatedWord) {
    lines.push(feedbackLine("warn", "Doppeltes Wortmuster erkannt. Noch einmal gegenlesen."));
  }

  if (style.lowercaseStart) {
    lines.push(feedbackLine("warn", "Mindestens ein Satz beginnt vermutlich mit kleinem Buchstaben."));
  }

  lines.push(feedbackLine("ok", "Browser-Rechtschreibprüfung ist aktiv."));

  const hardError = lines.some((line) => line.includes("error"));
  return { ok: !hardError, lines };
}

function validateSubtitle(value) {
  const result = validateTextField(value, { maxChars: 80, required: true, minChars: 25 });
  const lower = value.toLowerCase();
  const firstNames = getFirstNames();
  const themeKeywords = getThemeKeywords();

  if (firstNames.length) {
    const missingNames = firstNames.filter((name) => !lower.includes(name.toLowerCase()));
    if (missingNames.length) {
      result.lines.push(feedbackLine("warn", `Vornamen fehlen im Untertitel: ${missingNames.join(", ")}.`));
      result.ok = false;
    } else {
      result.lines.push(feedbackLine("ok", "Vornamen der Debattierenden sind erkennbar."));
    }
  }

  if (themeKeywords.length) {
    const matches = themeKeywords.filter((keyword) => lower.includes(keyword));
    if (matches.length === 0) {
      result.lines.push(feedbackLine("warn", "Debattenthema ist im Untertitel nicht klar erkennbar."));
      result.ok = false;
    } else {
      result.lines.push(feedbackLine("ok", "Debattenthema ist im Untertitel verankert."));
    }
  }

  return result;
}

function validateLead(value) {
  const result = validateTextField(value, { maxChars: 500, required: true, minChars: 180 });
  const coverage = analyzeLead(value);
  const missing = Object.entries(coverage)
    .filter(([, present]) => !present)
    .map(([key]) => key.toUpperCase());

  if (missing.length) {
    result.lines.push(feedbackLine("warn", `W-Fragen/Bilanz noch unklar: ${missing.join(", ")}.`));
    result.ok = false;
  } else {
    result.lines.push(feedbackLine("ok", "Wer, Was, Wann, Wo, Warum und Bilanz wurden erkannt."));
  }

  return result;
}

function validateAll() {
  const results = {};
  results.authorNames = validateTextField(state.authorNames, {
    maxChars: 100,
    required: true,
    minChars: 8
  });
  results.title = validateTextField(state.title, { maxChars: 50, required: true, minChars: 12 });
  results.subtitle = validateSubtitle(state.subtitle);
  results.lead = validateLead(state.lead);
  results.blockOne = validateTextField(state.blockOne, {
    maxChars: 3000,
    required: true,
    minChars: 2200
  });
  results.blockTwo = validateTextField(state.blockTwo, {
    maxChars: 3000,
    required: true,
    minChars: 2200
  });
  results.debateCaption = validateTextField(state.debateCaption, {
    maxChars: 100,
    required: false,
    minChars: 12
  });
  results.quoteText = validateTextField(state.quoteText, {
    maxChars: 220,
    required: true,
    minChars: 25
  });
  results.quoteSpeaker = validateTextField(state.quoteSpeaker, {
    maxChars: 80,
    required: true,
    minChars: 3
  });

  Object.entries(results).forEach(([field, result]) => setFeedback(field, result.lines));
  return results;
}

function updateReadiness(results) {
  const items = [];
  const missingRequired = requiredFields.filter((field) => !String(state[field] || "").trim());
  if (missingRequired.length) {
    items.push({
      level: "error",
      label: `Pflichtfelder fehlen: ${missingRequired.join(", ")}.`
    });
  } else {
    items.push({
      level: "ok",
      label: "Alle Pflichtfelder sind befüllt."
    });
  }

  const invalidFields = Object.entries(results)
    .filter(([, result]) => !result.ok)
    .map(([field]) => field);

  if (invalidFields.length) {
    items.push({
      level: "warn",
      label: `Formale oder stilistische Hinweise offen: ${invalidFields.join(", ")}.`
    });
  } else {
    items.push({
      level: "ok",
      label: "Formale Kriterien und Stilprüfungen sehen gut aus."
    });
  }

  if (!state.debateImageDataUrl) {
    items.push({
      level: "warn",
      label: "Debattenbild fehlt noch."
    });
  } else {
    items.push({
      level: "ok",
      label: "Debattenbild ist eingebunden."
    });
  }

  if (!state.authorImageDataUrl) {
    items.push({
      level: "warn",
      label: "Autorenbild fehlt noch."
    });
  } else {
    items.push({
      level: "ok",
      label: "Autorenbild ist eingebunden."
    });
  }

  readinessList.innerHTML = items
    .map((item) => `<li class="${item.level}">${escapeHtml(item.label)}</li>`)
    .join("");

  exportButton.disabled = missingRequired.length > 0;
}

function updatePreviewImages() {
  if (state.authorImageDataUrl) {
    preview.authorImage.src = state.authorImageDataUrl;
    preview.authorImage.hidden = false;
    preview.authorPlaceholder.hidden = true;
  } else {
    preview.authorImage.removeAttribute("src");
    preview.authorImage.hidden = true;
    preview.authorPlaceholder.hidden = false;
  }

  if (state.debateImageDataUrl) {
    preview.debateImage.src = state.debateImageDataUrl;
    preview.debateImage.hidden = false;
    preview.debatePlaceholder.hidden = true;
  } else {
    preview.debateImage.removeAttribute("src");
    preview.debateImage.hidden = true;
    preview.debatePlaceholder.hidden = false;
  }
}

function renderPreview() {
  preview.authorsInline.textContent = state.authorNames || "Autor*innenteam";
  preview.dateInline.textContent = formatDisplayDate(state.reportDate);
  preview.dateBadge.textContent = formatDisplayDate(state.reportDate);
  preview.theme.textContent = state.debateTheme || "Debattenthema";
  preview.title.textContent = state.title || "Packender, zur Debatte passender Titel";
  preview.subtitle.textContent =
    state.subtitle || "Untertitel mit Fragestellung und Vornamen der Debattierenden";
  preview.lead.textContent =
    state.lead || "Zugkraeftiger Lead mit W-Fragen, Kontext und einer kurzen Bilanz der Debatte.";
  preview.debaters.textContent = state.debaterNames || "Debattierende";
  preview.quote.textContent = state.quoteText ? `„${state.quoteText}“` : "„Knackiges Zitat aus der Debatte“";
  preview.quoteSpeaker.textContent = state.quoteSpeaker || "Name Debattierende:r";
  preview.authorCaption.textContent = state.authorNames || "Autor*innennamen";
  preview.debateCaption.textContent = state.debateCaption || "Bildlegende des Debattenbilds";
  preview.blockOne.innerHTML = paragraphsToHtml(
    state.blockOne,
    "Bericht Teil 1: Einfuehrungsreden und freie Aussprache."
  );
  preview.blockTwo.innerHTML = paragraphsToHtml(
    state.blockTwo,
    "Bericht Teil 2: Publikumsbeteiligung, Schluss und Mentimeter-Resultate."
  );
  updatePreviewImages();
}

async function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Datei konnte nicht gelesen werden."));
    reader.readAsDataURL(file);
  });
}

function textToDocxRuns(text, options = {}) {
  const lines = String(text || "").split("\n");
  const runProps = [];

  if (options.bold) runProps.push("<w:b/>");
  if (options.italic) runProps.push("<w:i/>");
  if (options.color) runProps.push(`<w:color w:val="${options.color}"/>`);
  if (options.fontSize) runProps.push(`<w:sz w:val="${options.fontSize}"/><w:szCs w:val="${options.fontSize}"/>`);
  if (options.font) {
    runProps.push(
      `<w:rFonts w:ascii="${escapeXml(options.font)}" w:hAnsi="${escapeXml(options.font)}" w:cs="${escapeXml(options.font)}"/>`
    );
  }

  const runPropXml = runProps.length ? `<w:rPr>${runProps.join("")}</w:rPr>` : "";
  return lines
    .map((line, index) => {
      const breakNode = index < lines.length - 1 ? "<w:br/>" : "";
      return `<w:r>${runPropXml}<w:t xml:space="preserve">${escapeXml(line)}</w:t>${breakNode}</w:r>`;
    })
    .join("");
}

function createParagraph(text, options = {}) {
  const pProps = [];
  const spacing = [];

  if (options.align) pProps.push(`<w:jc w:val="${options.align}"/>`);
  if (options.keepNext) pProps.push("<w:keepNext/>");
  if (options.pageBreakBefore) pProps.push("<w:pageBreakBefore/>");
  if (options.spacingBefore !== undefined) spacing.push(`w:before="${options.spacingBefore}"`);
  if (options.spacingAfter !== undefined) spacing.push(`w:after="${options.spacingAfter}"`);
  if (options.line) spacing.push(`w:line="${options.line}" w:lineRule="auto"`);
  if (spacing.length) pProps.push(`<w:spacing ${spacing.join(" ")}/>`);
  if (options.borderBottom) {
    pProps.push(
      `<w:pBdr><w:bottom w:val="single" w:sz="${options.borderBottom.size || 12}" w:space="1" w:color="${options.borderBottom.color || "111111"}"/></w:pBdr>`
    );
  }

  const pPr = pProps.length ? `<w:pPr>${pProps.join("")}</w:pPr>` : "";
  const runs = options.rawRuns || textToDocxRuns(text, options.run || {});
  return `<w:p>${pPr}${runs}</w:p>`;
}

function createParagraphsFromText(text, options = {}) {
  const paragraphs = String(text || "")
    .split(/\n{2,}/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (!paragraphs.length) {
    return createParagraph(options.placeholder || "", {
      ...(options.paragraph || {}),
      run: options.run || {}
    });
  }

  return paragraphs
    .map((paragraphText) =>
      createParagraph(paragraphText, {
        ...(options.paragraph || {}),
        run: options.run || {}
      })
    )
    .join("");
}

function createTableCell(innerXml, options = {}) {
  const tcProps = [
    `<w:tcW w:w="${options.width || 2400}" w:type="dxa"/>`,
    "<w:vAlign w:val=\"top\"/>",
    "<w:tcMar><w:top w:w=\"60\" w:type=\"dxa\"/><w:left w:w=\"90\" w:type=\"dxa\"/><w:bottom w:w=\"60\" w:type=\"dxa\"/><w:right w:w=\"90\" w:type=\"dxa\"/></w:tcMar>"
  ];

  if (options.border) {
    tcProps.push(
      `<w:tcBorders><w:top w:val="single" w:sz="${options.border}" w:color="111111"/><w:left w:val="single" w:sz="${options.border}" w:color="111111"/><w:bottom w:val="single" w:sz="${options.border}" w:color="111111"/><w:right w:val="single" w:sz="${options.border}" w:color="111111"/></w:tcBorders>`
    );
  } else {
    tcProps.push(
      "<w:tcBorders><w:top w:val=\"nil\"/><w:left w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/></w:tcBorders>"
    );
  }

  if (options.shading) {
    tcProps.push(`<w:shd w:val="clear" w:color="auto" w:fill="${options.shading}"/>`);
  }

  return `<w:tc><w:tcPr>${tcProps.join("")}</w:tcPr>${innerXml}</w:tc>`;
}

function createTable(rowsXml, options = {}) {
  const tableProps = [
    `<w:tblW w:w="${options.width || 0}" w:type="${options.width ? "dxa" : "auto"}"/>`,
    "<w:tblBorders><w:top w:val=\"nil\"/><w:left w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/><w:insideH w:val=\"nil\"/><w:insideV w:val=\"nil\"/></w:tblBorders>"
  ];

  if (options.layoutFixed) {
    tableProps.push("<w:tblLayout w:type=\"fixed\"/>");
  }

  return `<w:tbl><w:tblPr>${tableProps.join("")}</w:tblPr>${rowsXml}</w:tbl>`;
}

function createTableRow(cellsXml) {
  return `<w:tr>${cellsXml}</w:tr>`;
}

function inchesToEmu(value) {
  return Math.round(value * 914400);
}

function dataUrlToBinary(dataUrl) {
  const match = /^data:([^;]+);base64,(.+)$/s.exec(String(dataUrl || ""));
  if (!match) return null;

  const mimeType = match[1];
  const binaryString = atob(match[2]);
  const bytes = new Uint8Array(binaryString.length);
  for (let index = 0; index < binaryString.length; index += 1) {
    bytes[index] = binaryString.charCodeAt(index);
  }

  const extensionMap = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/gif": "gif",
    "image/webp": "webp"
  };

  return {
    mimeType,
    extension: extensionMap[mimeType] || "bin",
    bytes
  };
}

function createImageParagraph(media, options = {}) {
  if (!media) {
    return createParagraph(options.placeholder || "Bild", {
      align: "center",
      spacingAfter: options.spacingAfter !== undefined ? options.spacingAfter : 80,
      run: { font: "Arial", fontSize: 18, color: "666666" }
    });
  }

  const cx = options.cx || inchesToEmu(2.4);
  const cy = options.cy || inchesToEmu(3.0);
  const docPrId = options.docPrId || 1;
  const runs =
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${docPrId}" name="${escapeXml(media.name)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${docPrId}" name="${escapeXml(media.name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${media.relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

  return createParagraph("", {
    align: options.align || "center",
    spacingAfter: options.spacingAfter !== undefined ? options.spacingAfter : 80,
    rawRuns: runs
  });
}

function makeTextBytes(text) {
  return new TextEncoder().encode(text);
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let current = i;
    for (let bit = 0; bit < 8; bit += 1) {
      current = (current & 1) ? (0xedb88320 ^ (current >>> 1)) : (current >>> 1);
    }
    table[i] = current >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dateToDosParts(date) {
  const year = Math.max(1980, date.getFullYear());
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = Math.floor(date.getSeconds() / 2);
  return {
    dosTime: (hours << 11) | (minutes << 5) | seconds,
    dosDate: ((year - 1980) << 9) | (month << 5) | day
  };
}

function concatUint8Arrays(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

function createStoredZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const { dosTime, dosDate } = dateToDosParts(new Date());

  entries.forEach((entry) => {
    const nameBytes = makeTextBytes(entry.name);
    const data = entry.data;
    const crc = crc32(data);
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    offset += localHeader.length + data.length;
  });

  const centralDirectory = concatUint8Arrays(centralParts);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralDirectory.length, true);
  endView.setUint32(16, offset, true);

  return new Blob([concatUint8Arrays([...localParts, centralDirectory, endRecord])], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

function buildDocx() {
  const dateLabel = formatDisplayDate(state.reportDate);
  const isoDate = new Date().toISOString();
  const titleText = state.title || "Packender, zur Debatte passender Titel";
  const subtitleText = state.subtitle || "Untertitel mit Fragestellung und Vornamen der Debattierenden";
  const leadText = state.lead || "Zugkraeftiger Lead mit W-Fragen, Kontext und einer kurzen Bilanz der Debatte.";
  const quoteText = state.quoteText || "Knackiges Zitat aus der Debatte";
  const quoteSpeaker = state.quoteSpeaker || "Name Debattierende:r";
  const authorNames = state.authorNames || "Autor*innenteam";
  const debateTheme = state.debateTheme || "Debattenthema";
  const debateCaption = state.debateCaption || "Bildlegende des Debattenbilds";
  const debaterNames = state.debaterNames || "Debattierende";

  const mediaEntries = [];
  let nextRelId = 3;
  let nextDocPrId = 10;

  function registerImage(dataUrl, fallbackName) {
    const parsed = dataUrlToBinary(dataUrl);
    if (!parsed) return null;
    const mediaIndex = mediaEntries.length + 1;
    const target = `media/${fallbackName}-${mediaIndex}.${parsed.extension}`;
    const media = {
      relId: `rId${nextRelId}`,
      target,
      name: `${fallbackName}-${mediaIndex}.${parsed.extension}`,
      mimeType: parsed.mimeType,
      extension: parsed.extension,
      bytes: parsed.bytes,
      docPrId: nextDocPrId
    };
    nextRelId += 1;
    nextDocPrId += 1;
    mediaEntries.push(media);
    return media;
  }

  const authorImage = registerImage(state.authorImageDataUrl, "autor");
  const debateImage = registerImage(state.debateImageDataUrl, "debatte");

  const mastheadTable = createTable(
    createTableRow(
      createTableCell(
        createParagraph("KSA", {
          run: { font: "Arial Black", fontSize: 34, color: "DF1B12" },
          spacingAfter: 30
        }) +
          createParagraph("Kantonsschule Ausserschwyz", {
            run: { font: "Arial", fontSize: 16 },
            spacingAfter: 20
          }) +
          createParagraph("Gymi und FMS", {
            run: { font: "Arial", fontSize: 16 },
            spacingAfter: 0
          }),
        { width: 6400 }
      ) +
        createTableCell(
          createParagraph(`Bericht von ${authorNames}`, {
            align: "right",
            run: { font: "Georgia", fontSize: 18 },
            spacingAfter: 20
          }) +
            createParagraph(dateLabel, {
              align: "right",
              run: { font: "Georgia", fontSize: 18 },
              spacingAfter: 0
            }),
          { width: 3200 }
        )
    ),
    { width: 9600, layoutFixed: true }
  );

  const sectionBar = createParagraph("DEBATTENBERICHT", {
    align: "center",
    spacingBefore: 120,
    spacingAfter: 140,
    borderBottom: { size: 18, color: "111111" },
    run: { font: "Arial Black", fontSize: 24 }
  });

  const titleTable = createTable(
    createTableRow(
      createTableCell(
        createParagraph(debateTheme, {
          run: { font: "Georgia", fontSize: 22 },
          spacingAfter: 40
        }) +
          createParagraph(titleText, {
            run: { font: "Arial Black", fontSize: 44, color: "284D92" },
            spacingAfter: 70
          }) +
          createParagraph(subtitleText, {
            run: { font: "Georgia", fontSize: 30 },
            spacingAfter: 20
          }),
        { width: 9600, border: 16, shading: "E1DDD7" }
      )
    ),
    { width: 9600, layoutFixed: true }
  );

  const authorCellContent =
    createImageParagraph(authorImage, {
      cx: inchesToEmu(2.2),
      cy: inchesToEmu(2.8),
      docPrId: authorImage ? authorImage.docPrId : 11,
      placeholder: "Autorenbild",
      spacingAfter: 50
    }) +
    createParagraph(authorNames, {
      run: { font: "Georgia", fontSize: 16 },
      spacingAfter: 0
    });

  const leadCellContent =
    createParagraph(leadText, {
      run: { font: "Georgia", fontSize: 25 },
      line: 340,
      spacingAfter: 120
    }) +
    createParagraph(`${debaterNames} | ${dateLabel}`, {
      run: { font: "Arial", fontSize: 14 },
      spacingAfter: 120
    }) +
    createParagraph(`"${quoteText}"`, {
      run: { font: "Arial Black", fontSize: 40, color: "284D92" },
      line: 390,
      spacingAfter: 40
    }) +
    createParagraph(`- ${quoteSpeaker}`, {
      align: "right",
      run: { font: "Georgia", fontSize: 18 },
      spacingAfter: 0
    });

  const ledeTable = createTable(
    createTableRow(
      createTableCell(authorCellContent, { width: 2800 }) +
        createTableCell(leadCellContent, { width: 6800, shading: "E6E2DB" })
    ),
    { width: 9600, layoutFixed: true }
  );

  const copyTable = createTable(
    createTableRow(
      createTableCell(
          createParagraph("Bericht Teil 1", {
            run: { font: "Arial Black", fontSize: 20 },
            spacingAfter: 60
          }) +
          createParagraphsFromText(state.blockOne, {
            placeholder: "Bericht Teil 1: Einfuehrungsreden und freie Aussprache.",
            paragraph: { spacingAfter: 110, line: 300 },
            run: { font: "Georgia", fontSize: 19 }
          }),
        { width: 4700, border: 6 }
      ) +
        createTableCell(
          createParagraph("Bericht Teil 2", {
            run: { font: "Arial Black", fontSize: 20 },
            spacingAfter: 60
          }) +
            createParagraphsFromText(state.blockTwo, {
              placeholder: "Bericht Teil 2: Publikumsbeteiligung, Schluss und Mentimeter-Resultate.",
              paragraph: { spacingAfter: 110, line: 300 },
              run: { font: "Georgia", fontSize: 19 }
            }),
          { width: 4900, border: 6 }
        )
    ),
    { width: 9600, layoutFixed: true }
  );

  const debateFigure =
    createImageParagraph(debateImage, {
      cx: inchesToEmu(6.6),
      cy: inchesToEmu(3.7),
      docPrId: debateImage ? debateImage.docPrId : 12,
      placeholder: "Debattenbild",
      spacingAfter: 40
    }) +
    createParagraph(debateCaption, {
      run: { font: "Georgia", fontSize: 16 },
      spacingAfter: 0
    });

  const bodyXml =
    mastheadTable +
    sectionBar +
    titleTable +
    createParagraph("", { spacingAfter: 80 }) +
    ledeTable +
    createParagraph("", { spacingAfter: 120 }) +
    copyTable +
    createParagraph("", { spacingAfter: 120 }) +
    debateFigure +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="900" w:right="850" w:bottom="900" w:left="850" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>`;

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" mc:Ignorable="w14 wp14"><w:body>${bodyXml}</w:body></w:document>`;

  const stylesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/><w:lang w:val="de-CH"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

  const settingsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>`;

  const appXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company>Kantonsschule Ausserschwyz</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>1.0</AppVersion></Properties>`;

  const coreXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(titleText)}</dc:title><dc:creator>${escapeXml(authorNames)}</dc:creator><cp:lastModifiedBy>${escapeXml(authorNames)}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:modified></cp:coreProperties>`;

  const contentTypeDefaults = new Map([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"]
  ]);
  mediaEntries.forEach((media) => {
    contentTypeDefaults.set(media.extension, media.mimeType);
  });

  const contentTypesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    Array.from(contentTypeDefaults.entries())
      .map(([extension, mimeType]) => `<Default Extension="${extension}" ContentType="${mimeType}"/>`)
      .join("") +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `</Types>`;

  const rootRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;

  const documentRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
    mediaEntries
      .map(
        (media) =>
          `<Relationship Id="${media.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${media.target}"/>`
      )
      .join("") +
    `</Relationships>`;

  const zipEntries = [
    { name: "[Content_Types].xml", data: makeTextBytes(contentTypesXml) },
    { name: "_rels/.rels", data: makeTextBytes(rootRelsXml) },
    { name: "docProps/app.xml", data: makeTextBytes(appXml) },
    { name: "docProps/core.xml", data: makeTextBytes(coreXml) },
    { name: "word/document.xml", data: makeTextBytes(documentXml) },
    { name: "word/styles.xml", data: makeTextBytes(stylesXml) },
    { name: "word/settings.xml", data: makeTextBytes(settingsXml) },
    { name: "word/_rels/document.xml.rels", data: makeTextBytes(documentRelsXml) }
  ];

  mediaEntries.forEach((media) => {
    zipEntries.push({ name: `word/${media.target}`, data: media.bytes });
  });

  return createStoredZip(zipEntries);
}

function downloadWordDocument() {
  const results = validateAll();
  updateReadiness(results);
  const missingRequired = requiredFields.filter((field) => !String(state[field] || "").trim());
  if (missingRequired.length) {
    window.alert("Pflichtfelder fehlen noch. Bitte zuerst alle Kernrubriken ausfuellen.");
    return;
  }

  const blob = buildDocx();
  const link = document.createElement("a");
  const datePart = state.reportDate || toInputDate(new Date());
  const objectUrl = URL.createObjectURL(blob);
  link.href = objectUrl;
  link.download = `debattenbericht-${datePart}.docx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

async function handleFileInput(input) {
  const file = input.files && input.files[0];
  if (!file) {
    return;
  }

  if (!/^image\/(png|jpeg)$/i.test(file.type)) {
    window.alert("Bitte ein PNG- oder JPG-Bild waehlen. Diese Formate werden im DOCX-Export unterstuetzt.");
    input.value = "";
    return;
  }

  const dataUrl = await fileToDataUrl(file);
  if (input.name === "authorImage") {
    state.authorImageDataUrl = dataUrl;
  }
  if (input.name === "debateImage") {
    state.debateImageDataUrl = dataUrl;
  }
  renderPreview();
  updateReadiness(validateAll());
}

function syncStateFromForm() {
  const formData = new FormData(form);
  Object.keys(defaultState).forEach((key) => {
    if (key.endsWith("DataUrl")) return;
    state[key] = String(formData.get(key) || "");
  });
}

function fillSampleData() {
  const sample = {
    reportDate: "2026-02-21",
    debateTheme: "Impfpflicht fuer Basisimpfungen",
    debaterNames: "Leonie, Anne, Nicolas, Michelle",
    authorNames: "Anne Krienbuehl, Leonie Moser",
    title: "Impfpflicht fuer Basisimpfungen?",
    subtitle: "Zur Impfpflicht debattieren Leonie, Anne, Nicolas und Michelle",
    lead:
      "Wer an der Kantonsschule Ausserschwyz ueber eine Impfpflicht debattierte, stritt am Freitag in einer pointierten Runde ueber Freiheit, Schutz und Verantwortung. Die Debattierenden lieferten klare Linien, das Publikum fragte scharf nach, und am Ende blieb eine knappe Bilanz: Die Pflicht spaltet, zwingt aber zur konkreten Abwaegung.",
    blockOne:
      "Zu Beginn setzten die Einfuehrungsreden den Ton der Debatte. Die Befuerworterseite rueckte den Schutz der Gesellschaft in den Mittelpunkt und verwies auf medizinische Risiken, die sich mit hohen Impfquoten begrenzen liessen. Die Gegenposition konterte frueh mit dem Hinweis auf die persoenliche Entscheidungsfreiheit und darauf, dass staatlicher Zwang in die koerperliche Selbstbestimmung eingreife.\n\nIn der freien Aussprache verdichteten sich diese Linien. Beide Seiten arbeiteten mit Beispielen aus dem Schulalltag, aus Spitaelern und aus der Frage, welche Verantwortung Einzelne gegenueber besonders gefaehrdeten Menschen tragen. Mehrfach wurde deutlich, dass die Kontroverse nicht an einzelnen Fakten hing, sondern an der Gewichtung von Sicherheit, Freiheit und Vertrauen in staatliche Eingriffe.\n\nBemerkenswert war, wie rasch sich die Debatte von Schlagworten loeste. Die Debattierenden praezisierten Begriffe, fragten nach Grenzen einer Pflicht und stellten wiederholt die Verhaeltnismaessigkeit ins Zentrum. Dadurch gewann der erste Teil an Schaerfe und blieb zugleich gut nachvollziehbar.",
    blockTwo:
      "Im zweiten Teil oeffnete sich die Debatte fuer das Publikum. Die Fragen zielten vor allem auf moegliche Alternativen zur Pflicht, auf Aufklaerungskampagnen und auf die Rolle von Schulen und Gesundheitseinrichtungen. Mehrere Wortmeldungen machten deutlich, dass ein bloss moralischer Appell vielen zu wenig weit geht, waehrend andere vor einer Eskalation durch Zwang warnten.\n\nDie Mentimeter-Resultate brachten eine nuancierte Zwischenbilanz. Zwar zeigte sich eine leichte Tendenz zugunsten staerkerer Massnahmen, doch das Bild blieb uneinheitlich. Genau diese Spannung praegte auch das Schlusssegment: Die Befuerwortenden argumentierten mit Praevention und Solidaritaet, die Gegenseite hielt an der individuellen Abwaegung fest.\n\nIm Schlusspunkt ueberzeugte vor allem, dass die Runde nicht bei Lagerdenken stehen blieb. Die Debatte endete offen, aber nicht beliebig: Sie zeigte, wie politisch aufgeladen das Thema bleibt und wie stark gute Argumente erst dann wirken, wenn sie konkret, aktiv und adressatenbezogen formuliert sind.",
    debateCaption: "Die Debatte an der KSA kreiste um Freiheit, Verantwortung und den Schutz vulnerabler Gruppen.",
    quoteText: "Anstatt die Leute zu zwingen, sollten wir Vertrauen schaffen.",
    quoteSpeaker: "Anne"
  };

  Object.entries(sample).forEach(([key, value]) => {
    const field = form.elements.namedItem(key);
    if (field) {
      field.value = value;
    }
    state[key] = value;
  });

  renderPreview();
  updateReadiness(validateAll());
}

function updateAll() {
  syncStateFromForm();
  renderPreview();
  updateReadiness(validateAll());
}

form.addEventListener("input", updateAll);
form.addEventListener("change", async (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    updateAll();
    return;
  }

  if (target.type === "file") {
    await handleFileInput(target);
    return;
  }

  updateAll();
});

sampleButton.addEventListener("click", fillSampleData);
exportButton.addEventListener("click", downloadWordDocument);

updateAll();
