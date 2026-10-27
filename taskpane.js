/* global Office, Word */
(() => {
  "use strict";

  const $ = id => document.getElementById(id);
  const STORAGE_KEY = "word-frame-studio.library.v2";

  const FONTS = ["Calibri", "Arial", "Georgia", "Times New Roman"];
  const MODES = ["card", "frame", "art"];
  const LINES = ["single", "double", "heavy", "round", "ascii"];
  const BORDERS = ["single", "double", "dashed", "dotted"];

  const SHAPES = {
    Star: [
      "    *",
      "   ***",
      "*********",
      "  *****",
      " **   **"
    ].join("\n"),
    Heart: [
      " **   **",
      "**** ****",
      "*********",
      " *******",
      "  *****",
      "   ***",
      "    *"
    ].join("\n"),
    Arrow: [
      "          /|",
      "---------/ |",
      "---------\\ |",
      "          \\|"
    ].join("\n"),
    Mountain: [
      "         /\\",
      "    /\\  /  \\",
      "   /  \\/    \\",
      "  /          \\",
      " /____________\\"
    ].join("\n"),
    Rocket: [
      "     /\\",
      "    /  \\",
      "   | () |",
      "   |    |",
      "  /|    |\\",
      " /_|____|_\\",
      "    /\\/\\"
    ].join("\n"),
    Robot: [
      "   .-------.",
      "   | o   o |",
      "   |  ---  |",
      "   '-------'",
      "  /|       |\\",
      " / |_______| \\",
      "     |   |",
      "    _|   |_"
    ].join("\n"),
    Castle: [
      "  |>>>          |>>>",
      "  |             |",
      " _|_ _ _     _ _|_",
      "| |_| |_|___| |_| |",
      "|                 |",
      "|   []   ___  []  |",
      "|       |   |     |",
      "|_______|___|_____|"
    ].join("\n"),
    Divider: "* - * - * - * - * - * - * - *",
    Checklist: "[ ] First item\n[ ] Second item\n[ ] Third item"
  };

  const TEMPLATES = {
    note: {
      title: "Project note",
      body: "Add a clear summary here.\n\nNext step: describe the action to take.",
      accent: "#2563eb",
      fill: "#eff6ff",
      border: "single",
      align: "left",
      font: "Calibri"
    },
    warning: {
      title: "Important",
      body: "Explain what needs attention.\n\nInclude the deadline or required action.",
      accent: "#b45309",
      fill: "#fffbeb",
      border: "single",
      align: "left",
      font: "Calibri"
    },
    quote: {
      title: "A thought worth sharing",
      body: "Write your quotation here.\n\n— Author",
      accent: "#64748b",
      fill: "#f8fafc",
      border: "single",
      align: "left",
      font: "Georgia"
    },
    certificate: {
      title: "Certificate of appreciation",
      body: "Presented to\n\nRecipient name\n\nIn recognition of outstanding contribution.\n\nDate • Signature",
      accent: "#92702a",
      fill: "#fffdf5",
      border: "double",
      align: "center",
      font: "Georgia"
    }
  };

  const DEFAULT = {
    mode: "card",
    ...TEMPLATES.note,
    size: 11,
    ink: "#243247",
    weight: 1,
    line: "single",
    width: 42,
    padding: 1,
    blankRows: 0,
    artText: SHAPES.Star
  };

  const FIELD_IDS = Object.keys(DEFAULT).filter(key => key !== "mode");

  let mode = "card";
  let wordReady = false;
  let busy = false;
  let library = [];

  function notify(message, error = false) {
    $("status").textContent = message;
    $("status").classList.toggle("error", error);
  }

  function normalizeText(value) {
    return String(value ?? "")
      .replace(/\r\n?/g, "\n")
      .replace(/\t/g, "    ")
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
  }

  function numberValue(value, min, max, fallback, step = 1) {
    const n = Number(value);
    if (value === "" || !Number.isFinite(n)) return fallback;
    return Math.max(min, Math.min(max, Math.round(n / step) * step));
  }

  function choice(value, allowed, fallback) {
    return allowed.includes(value) ? value : fallback;
  }

  function color(value, fallback) {
    return /^#[0-9a-f]{6}$/i.test(value || "") ? value : fallback;
  }

  function sanitize(input) {
    if (!input || typeof input !== "object" || Array.isArray(input)) {
      throw new Error("The design must be an object.");
    }

    return {
      mode: choice(input.mode, MODES, DEFAULT.mode),
      title: normalizeText(input.title ?? DEFAULT.title).slice(0, 160),
      body: normalizeText(input.body ?? DEFAULT.body).slice(0, 12000),
      artText: normalizeText(input.artText ?? DEFAULT.artText).slice(0, 12000),
      font: choice(input.font, FONTS, DEFAULT.font),
      size: numberValue(input.size, 8, 28, DEFAULT.size, 0.5),
      align: choice(input.align, ["left", "center", "right"], "left"),
      ink: color(input.ink, DEFAULT.ink),
      accent: color(input.accent, DEFAULT.accent),
      fill: color(input.fill, DEFAULT.fill),
      border: choice(input.border, BORDERS, DEFAULT.border),
      weight: numberValue(input.weight, 0.5, 4, DEFAULT.weight, 0.5),
      line: choice(input.line, LINES, DEFAULT.line),
      width: numberValue(input.width, 16, 90, DEFAULT.width),
      padding: numberValue(input.padding, 0, 5, DEFAULT.padding),
      blankRows: numberValue(input.blankRows, 0, 20, DEFAULT.blankRows)
    };
  }

  function readDesign() {
    const data = { mode };
    for (const id of FIELD_IDS) data[id] = $(id).value;
    return sanitize(data);
  }

  function applyDesign(data) {
    const design = sanitize(data);
    mode = design.mode;
    for (const id of FIELD_IDS) $(id).value = design[id];
    render();
  }

  // Wrap at whitespace where possible; split long words without dropping text.
  function wrapText(text, width) {
    return normalizeText(text).split("\n").flatMap(line => {
      let remaining = Array.from(line);
      const rows = [];

      while (remaining.length > width) {
        let cut = width;
        for (let i = width; i > 0; i--) {
          if (remaining[i] === " ") {
            cut = i;
            break;
          }
        }
        rows.push(remaining.slice(0, cut).join(""));
        remaining = remaining.slice(cut);
        if (remaining[0] === " ") remaining.shift();
      }

      rows.push(remaining.join(""));
      return rows;
    });
  }

  function alignText(text, width, alignment) {
    const free = Math.max(0, width - Array.from(text).length);
    const left = alignment === "right"
      ? free
      : alignment === "center" ? Math.floor(free / 2) : 0;

    return " ".repeat(left) + text + " ".repeat(free - left);
  }

  function frameText(d) {
    const styles = {
      single: ["┌", "┐", "└", "┘", "─", "│"],
      double: ["╔", "╗", "╚", "╝", "═", "║"],
      heavy: ["┏", "┓", "┗", "┛", "━", "┃"],
      round: ["╭", "╮", "╰", "╯", "─", "│"],
      ascii: ["+", "+", "+", "+", "-", "|"]
    };

    const [tl, tr, bl, br, horizontal, vertical] = styles[d.line];
    const contentWidth = d.width - d.padding * 2;
    const padding = " ".repeat(d.padding);
    const content = [];

    if (d.title) content.push(...wrapText(d.title, contentWidth));
    if (d.title && d.body) content.push("");
    if (d.body) content.push(...wrapText(d.body, contentWidth));
    if (!content.length) content.push("");

    for (let i = 0; i < d.blankRows; i++) content.push("");

    return [
      tl + horizontal.repeat(d.width) + tr,
      ...content.map(line =>
        vertical + padding +
        alignText(line, contentWidth, d.align) +
        padding + vertical
      ),
      bl + horizontal.repeat(d.width) + br
    ].join("\n");
  }

  function plainText(d) {
    if (d.mode === "frame") return frameText(d);
    if (d.mode === "art") return d.artText;
    return [d.title, d.body].filter(Boolean).join("\n\n");
  }

  function syncButtons() {
    document.querySelectorAll("[data-word]").forEach(button => {
      button.disabled = busy || !wordReady;
    });
    $("insert").textContent = busy ? "Working…" : "Insert design";
  }

  function render() {
    const d = readDesign();
    const card = d.mode === "card";

    document.querySelectorAll("[data-mode]").forEach(button => {
      button.setAttribute("aria-pressed", String(button.dataset.mode === mode));
    });

    $("cardControls").hidden = !card;
    $("textControls").hidden = mode === "art";
    $("frameControls").hidden = mode !== "frame";
    $("artControls").hidden = mode !== "art";
    $("cardPreview").hidden = !card;
    $("textPreview").hidden = card;
    $("font").disabled = !card;
    $("align").disabled = mode === "art";
    $("fontHint").hidden = card;

    document.querySelectorAll(".card-only").forEach(element => {
      element.hidden = !card;
    });

    $("replaceHint").hidden = $("destination").value !== "replace";

    const text = plainText(d);

    if (card) {
      const preview = $("cardPreview");
      const cssBorder = {
        single: "solid",
        double: "double",
        dashed: "dashed",
        dotted: "dotted"
      }[d.border];

      preview.style.border = `${d.weight}pt ${cssBorder} ${d.accent}`;
      preview.style.backgroundColor = d.fill;
      preview.style.color = d.ink;
      preview.style.fontFamily = d.font;
      preview.style.fontSize = `${d.size}pt`;
      preview.style.textAlign = d.align;

      $("previewTitle").textContent = d.title;
      $("previewTitle").hidden = !d.title;
      $("previewBody").textContent = d.body;
    } else {
      $("textPreview").textContent = text;
      $("textPreview").style.color = d.ink;
      $("textPreview").style.fontSize = `${d.size}pt`;
    }

    $("metrics").textContent =
      `${Array.from(text).length} chars · ${text.split("\n").length} lines`;

    syncButtons();
  }

  function escapeXml(value) {
    return normalizeText(value).replace(/[&<>"']/g, ch => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&apos;"
    }[ch]));
  }

  function paragraph(text, d, title = false) {
    const size = Math.round((title ? d.size + 3 : d.size) * 2);

    return `<w:p>
      <w:pPr>
        <w:jc w:val="${d.align}"/>
        <w:spacing w:after="${title ? 160 : 80}"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="${escapeXml(d.font)}"
                    w:hAnsi="${escapeXml(d.font)}"
                    w:cs="${escapeXml(d.font)}"/>
          ${title ? "<w:b/>" : ""}
          <w:color w:val="${d.ink.slice(1)}"/>
          <w:sz w:val="${size}"/>
          <w:szCs w:val="${size}"/>
        </w:rPr>
        <w:t xml:space="preserve">${escapeXml(text)}</w:t>
      </w:r>
    </w:p>`;
  }

  function cardOoxml(d) {
    const border = ["top", "left", "bottom", "right"].map(side =>
      `<w:${side} w:val="${d.border}"
        w:sz="${Math.round(d.weight * 8)}"
        w:color="${d.accent.slice(1)}"/>`
    ).join("");

    const margins = ["top", "left", "bottom", "right"].map(side =>
      `<w:${side} w:w="200" w:type="dxa"/>`
    ).join("");

    const paragraphs = [
      ...(d.title ? d.title.split("\n").map(t => paragraph(t, d, true)) : []),
      ...d.body.split("\n").map(t => paragraph(t, d))
    ].join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/_rels/.rels"
        pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
              Target="word/document.xml"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/document.xml"
        pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:tbl>
                <w:tblPr>
                  <w:tblW w:w="5000" w:type="pct"/>
                  <w:tblBorders>${border}</w:tblBorders>
                  <w:tblCellMar>${margins}</w:tblCellMar>
                </w:tblPr>
                <w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
                <w:tr>
                  <w:tc>
                    <w:tcPr>
                      <w:tcW w:w="5000" w:type="pct"/>
                      <w:shd w:val="clear" w:color="auto"
                             w:fill="${d.fill.slice(1)}"/>
                    </w:tcPr>
                    ${paragraphs}
                  </w:tc>
                </w:tr>
              </w:tbl>
              <w:p/>
            </w:body>
          </w:document>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>`;
  }

  async function withWord(action, success) {
    if (!wordReady) {
      notify("Open this add-in inside a supported version of Microsoft Word.", true);
      return;
    }

    if (busy) return;
    busy = true;
    syncButtons();

    try {
      await Word.run(action);
      notify(success);
    } catch (error) {
      const code = error.code ? ` (${error.code})` : "";
      notify(
        `Word could not complete the operation${code}. ` +
        `${error.message || "Check that the document is editable and try again."}`,
        true
      );
    } finally {
      busy = false;
      syncButtons();
    }
  }

  async function insertDesign() {
    const d = readDesign();
    const text = plainText(d);
    const destination = $("destination").value;

    if (!text.trim()) {
      notify("Add some text before inserting.", true);
      return;
    }

    await withWord(async context => {
      const target = destination === "end"
        ? context.document.body.getRange("End")
        : context.document.getSelection();

      const location = destination === "replace" ? "Replace" : "End";

      if (d.mode === "card") {
        target.insertOoxml(cardOoxml(d), location);
      } else {
        // Format the returned inserted range, not the original selection.
        const inserted = target.insertText("\n" + text + "\n", location);
        inserted.font.name = "Consolas";
        inserted.font.size = d.size;
        inserted.font.color = d.ink;

        const paragraphs = inserted.paragraphFormat;
        paragraphs.spaceBefore = 0;
        paragraphs.spaceAfter = 0;
        paragraphs.lineSpacing = d.size * 1.2;
        paragraphs.alignment = "Left";
      }

      await context.sync();
    }, "Design inserted. You can edit it directly in Word.");
  }

  async function useSelection() {
    await withWord(async context => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text.trim()) {
        throw new Error("Select some text in Word first.");
      }
      if (selection.text.length > 12000) {
        throw new Error("Select 12,000 characters or fewer.");
      }

      $("body").value = normalizeText(selection.text);
      render();
    }, "Selected text loaded. The document has not been changed.");
  }

  async function copyText() {
    try {
      if (!navigator.clipboard || !navigator.clipboard.writeText) {
        throw new Error("Clipboard unavailable");
      }
      await navigator.clipboard.writeText(plainText(readDesign()));
      notify("Plain text copied.");
    } catch {
      notify(
        "Clipboard access was blocked. Select and copy the preview, or use Download .txt.",
        true
      );
    }
  }

  function downloadFile(name, content, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = name;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function validateLibrary(data) {
    if (!data || data.version !== 2 || !Array.isArray(data.designs)) {
      throw new Error("Choose a Frame Studio version 2 library.");
    }

    if (data.designs.length > 50) {
      throw new Error("A library can contain at most 50 designs.");
    }

    const seen = new Set();

    return data.designs.map(item => {
      if (!item || typeof item.name !== "string" || !item.name.trim()) {
        throw new Error("Every design needs a name.");
      }

      const name = item.name.trim().slice(0, 60);
      const key = name.toLowerCase();

      if (seen.has(key)) {
        throw new Error("The library contains duplicate design names.");
      }
      seen.add(key);

      return { name, design: sanitize(item.design) };
    });
  }

  function refreshLibrary(selectedName) {
    const select = $("saved");
    select.replaceChildren();

    if (!library.length) {
      select.add(new Option("No saved designs yet", ""));
    } else {
      library.forEach((item, index) => {
        select.add(new Option(item.name, String(index)));
        if (item.name === selectedName) select.value = String(index);
      });
    }

    $("load").disabled = !library.length;
    $("delete").disabled = !library.length;
  }

  function commitLibrary(next, selectedName) {
    // Only update in-memory state after storage succeeds.
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ version: 2, designs: next })
    );
    library = next;
    refreshLibrary(selectedName);
  }

  function saveDesign() {
    const name = $("designName").value.trim().slice(0, 60);

    if (!name) {
      notify("Enter a name for your design.", true);
      $("designName").focus();
      return;
    }

    const existing = library.findIndex(
      item => item.name.toLowerCase() === name.toLowerCase()
    );

    if (existing >= 0) {
      notify("That name is already saved. Use a different name, or delete the old design.", true);
      return;
    }

    if (library.length >= 50) {
      notify("The library is full. Export a backup and delete a design first.", true);
      return;
    }

    try {
      commitLibrary([...library, { name, design: readDesign() }], name);
      notify("Design saved in this browser profile.");
    } catch {
      notify("Saving failed. Browser storage may be blocked or full.", true);
    }
  }

  async function importLibrary(file) {
    if (!file) return;

    try {
      if (file.size > 2 * 1024 * 1024) {
        throw new Error("Choose a JSON file smaller than 2 MB.");
      }

      const incoming = validateLibrary(JSON.parse(await file.text()));
      const names = new Set(library.map(item => item.name.toLowerCase()));
      const additions = [];

      for (const item of incoming) {
        let name = item.name;
        let suffix = 2;

        while (names.has(name.toLowerCase())) {
          name = `${item.name.slice(0, 50)} (${suffix++})`;
        }

        names.add(name.toLowerCase());
        additions.push({ ...item, name });
      }

      if (library.length + additions.length > 50) {
        throw new Error("The combined library would exceed 50 designs.");
      }

      commitLibrary([...library, ...additions]);
      notify(`Imported ${additions.length} designs. Existing designs were preserved.`);
    } catch (error) {
      notify(`Import failed: ${error.message}`, true);
    } finally {
      $("importFile").value = "";
    }
  }

  function initializeWord() {
    if (!window.Office) {
      $("connection").textContent = "Preview mode · Office.js unavailable";
      notify("You can preview designs. Open the add-in in Word to insert them.");
      return;
    }

    const timeout = setTimeout(() => {
      if (!wordReady) {
        $("connection").textContent = "Preview mode · Word not connected";
      }
    }, 10000);

    Office.onReady(info => {
      clearTimeout(timeout);

      wordReady =
        info.host === Office.HostType.Word &&
        !!window.Word &&
        Office.context.requirements.isSetSupported("WordApi", "1.3");

      $("connection").textContent = wordReady
        ? "Connected to Word"
        : "Preview mode · Requires WordApi 1.3";

      syncButtons();
    });
  }

  function initialize() {
    Object.keys(SHAPES).forEach(name => {
      $("shape").add(new Option(name, name));
    });

    document.querySelectorAll("[data-mode]").forEach(button => {
      button.addEventListener("click", () => {
        mode = button.dataset.mode;
        render();
      });
    });

    for (const id of FIELD_IDS) {
      $(id).addEventListener("input", render);
      $(id).addEventListener("change", render);
    }

    $("template").addEventListener("change", () => {
      applyDesign({
        ...readDesign(),
        ...TEMPLATES[$("template").value],
        mode: "card"
      });
    });

    $("shape").addEventListener("change", () => {
      $("artText").value = SHAPES[$("shape").value];
      render();
    });

    $("destination").addEventListener("change", render);
    $("insert").addEventListener("click", insertDesign);
    $("selection").addEventListener("click", useSelection);
    $("copy").addEventListener("click", copyText);

    $("download").addEventListener("click", () => {
      downloadFile(
        "frame-studio-design.txt",
        plainText(readDesign()),
        "text/plain;charset=utf-8"
      );
      notify("Download requested. If Word blocks it, open the task pane URL in your browser.");
    });

    $("reset").addEventListener("click", () => {
      $("template").value = "note";
      $("shape").value = "Star";
      applyDesign({ ...DEFAULT, mode });
      notify("Editor reset. Saved designs are unchanged.");
    });

    $("reload").addEventListener("click", () => {
      const url = new URL(window.location.href);
      url.searchParams.set("reload", Date.now().toString());
      window.location.replace(url.href);
    });

    $("save").addEventListener("click", saveDesign);

    $("load").addEventListener("click", () => {
      const item = library[Number($("saved").value)];
      if (!item) return;
      applyDesign(item.design);
      $("designName").value = item.name;
      notify(`Loaded “${item.name}”.`);
    });

    $("delete").addEventListener("click", () => {
      const index = Number($("saved").value);
      if (!library[index]) return;

      try {
        commitLibrary(library.filter((_, i) => i !== index));
        notify("Saved design deleted.");
      } catch {
        notify("Could not update browser storage.", true);
      }
    });

    $("export").addEventListener("click", () => {
      downloadFile(
        "frame-studio-library.json",
        JSON.stringify({ version: 2, designs: library }, null, 2),
        "application/json"
      );
      notify("Library export requested.");
    });

    $("import").addEventListener("click", () => $("importFile").click());
    $("importFile").addEventListener("change", event => {
      importLibrary(event.target.files[0]);
    });

    applyDesign(DEFAULT);

    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) library = validateLibrary(JSON.parse(stored));
    } catch {
      notify(
        "The saved library could not be read. Storage may be blocked or its data may be invalid.",
        true
      );
    }

    refreshLibrary();
    initializeWord();
  }

  initialize();
})();
