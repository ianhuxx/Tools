function cleanFileName(input) {
  return (input || "untitled")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsText(file);
  });
}

function md5HexFromBase64(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return SparkMD5.ArrayBuffer.hash(bytes.buffer);
}

function parseEnex(xmlText) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(xmlText, "application/xml");

  if (xml.querySelector("parsererror")) {
    throw new Error("Invalid ENEX XML");
  }

  const notes = Array.from(xml.querySelectorAll("note"));
  return notes.map((noteEl) => {
    const title = noteEl.querySelector("title")?.textContent?.trim() || "Untitled";
    const created = noteEl.querySelector("created")?.textContent?.trim() || "";
    const updated = noteEl.querySelector("updated")?.textContent?.trim() || "";

    const resources = {};
    for (const res of noteEl.querySelectorAll("resource")) {
      const dataEl = res.querySelector("data");
      const mime = res.querySelector("mime")?.textContent?.trim() || "application/octet-stream";
      let hash = res.querySelector("data")?.getAttribute("hash") || "";
      const fileName =
        res.querySelector("resource-attributes > file-name")?.textContent?.trim() || "attachment";
      const base64 = dataEl?.textContent?.replace(/\s+/g, "") || "";
      if (!hash && base64) {
        hash = md5HexFromBase64(base64);
      }
      if (hash && base64) {
        resources[hash.toLowerCase()] = { base64, mime, fileName };
      }
    }

    const contentRaw = noteEl.querySelector("content")?.textContent || "";
    const enml = contentRaw.replace(/^\s*<\?xml[^>]*>\s*<!DOCTYPE[^>]*>/i, "");

    return { title, created, updated, enml, resources };
  });
}

function buildDocHtml(note) {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = note.enml;

  const enNote = wrapper.querySelector("en-note") || wrapper;

  for (const media of enNote.querySelectorAll("en-media")) {
    const hash = (media.getAttribute("hash") || "").toLowerCase();
    const resource = note.resources[hash];

    if (resource && resource.mime.startsWith("image/")) {
      const img = document.createElement("img");
      img.src = `data:${resource.mime};base64,${resource.base64}`;
      img.alt = resource.fileName || "image";
      img.style.maxWidth = "100%";
      media.replaceWith(img);
    } else if (resource) {
      const p = document.createElement("p");
      p.textContent = `[Attachment: ${resource.fileName}]`;
      media.replaceWith(p);
    } else {
      const p = document.createElement("p");
      p.textContent = "[Embedded resource could not be mapped]";
      media.replaceWith(p);
    }
  }

  const created = note.created ? `<p><strong>Created:</strong> ${note.created}</p>` : "";
  const updated = note.updated ? `<p><strong>Updated:</strong> ${note.updated}</p>` : "";

  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body { font-family: Calibri, Arial, sans-serif; line-height: 1.4; }
      img { max-width: 100%; height: auto; }
      pre { white-space: pre-wrap; }
      table { border-collapse: collapse; }
      td, th { border: 1px solid #ccc; padding: 4px; }
    </style>
  </head>
  <body>
    <h1>${note.title.replace(/</g, "&lt;")}</h1>
    ${created}
    ${updated}
    <hr />
    ${enNote.innerHTML}
  </body>
</html>`;
}

function updateProgress(progressBarEl, progressTextEl, done, total, label) {
  const safeTotal = Math.max(total, 1);
  const percent = Math.round((done / safeTotal) * 100);
  progressBarEl.style.width = `${percent}%`;
  progressTextEl.textContent = label || `Progress: ${done}/${total} (${percent}%)`;
}

async function countTotalNotes(files) {
  let total = 0;
  for (const file of files) {
    const xmlText = await readFileAsText(file);
    total += parseEnex(xmlText).length;
  }
  return total;
}

async function convertFiles(files, controls) {
  const {
    statusEl,
    progressBarEl,
    progressTextEl,
    singleOutputEnabled,
    errorBoxEl
  } = controls;

  errorBoxEl.hidden = true;
  errorBoxEl.textContent = "";

  statusEl.textContent = "Scanning ENEX files...";
  const totalNotes = await countTotalNotes(files);

  if (totalNotes === 0) {
    throw new Error("No notes found in selected ENEX file(s).");
  }

  let convertedNotes = 0;
  const failures = [];

  updateProgress(progressBarEl, progressTextEl, 0, totalNotes, `Starting conversion (0/${totalNotes})...`);

  const zip = new JSZip();
  const flatOutputs = [];

  for (const file of files) {
    statusEl.textContent = `Reading ${file.name}...`;
    let xmlText;
    let notes;
    try {
      xmlText = await readFileAsText(file);
      notes = parseEnex(xmlText);
    } catch (err) {
      failures.push(`File ${file.name}: ${err.message || String(err)}`);
      continue;
    }

    const notebookFolder = zip.folder(cleanFileName(file.name.replace(/\.enex$/i, "")));

    for (let i = 0; i < notes.length; i += 1) {
      const note = notes[i];
      try {
        statusEl.textContent = `Converting: ${note.title}`;
        const html = buildDocHtml(note);
        const blob = window.htmlDocx.asBlob(html);
        const fileName = `${String(i + 1).padStart(4, "0")}-${cleanFileName(note.title)}.docx`;
        notebookFolder.file(fileName, blob);
        flatOutputs.push({ fileName, blob, notebook: cleanFileName(file.name.replace(/\.enex$/i, "")) });
        convertedNotes += 1;
      } catch (err) {
        failures.push(`Note "${note.title}" in ${file.name}: ${err.message || String(err)}`);
      }
      updateProgress(progressBarEl, progressTextEl, convertedNotes, totalNotes);
    }
  }

  if (convertedNotes === 0) {
    throw new Error(`Conversion failed for all notes. ${failures[0] || "Unknown error."}`);
  }

  statusEl.textContent = "Preparing download...";

  if (singleOutputEnabled && flatOutputs.length === 1) {
    saveAs(flatOutputs[0].blob, flatOutputs[0].fileName);
    statusEl.textContent = "Done. Downloaded 1 DOCX file.";
  } else {
    const finalZip = await zip.generateAsync({ type: "blob" });
    saveAs(finalZip, `enex-to-onenote-${Date.now()}.zip`);
    statusEl.textContent = `Done. Converted ${convertedNotes} note(s) into a structured ZIP.`;
  }

  if (failures.length > 0) {
    errorBoxEl.hidden = false;
    const max = failures.slice(0, 8);
    const more = failures.length > max.length ? `\n...and ${failures.length - max.length} more issue(s).` : "";
    errorBoxEl.textContent = `Some items could not be converted:\n- ${max.join("\n- ")}${more}`;
  }
}

function initDragAndDrop(dropZone, filesInput, onFilesSelected) {
  const prevent = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, prevent);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => dropZone.classList.add("drag-active"));
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, () => dropZone.classList.remove("drag-active"));
  });

  dropZone.addEventListener("drop", (event) => {
    const droppedFiles = Array.from(event.dataTransfer?.files || []).filter((f) =>
      /\.enex$/i.test(f.name)
    );
    const transfer = new DataTransfer();
    droppedFiles.forEach((f) => transfer.items.add(f));
    filesInput.files = transfer.files;
    onFilesSelected();
  });

  dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      filesInput.click();
    }
  });
}

function init() {
  const filesInput = document.getElementById("enexFiles");
  const convertBtn = document.getElementById("convertBtn");
  const statusEl = document.getElementById("status");
  const progressBarEl = document.getElementById("progressBar");
  const progressTextEl = document.getElementById("progressText");
  const singleOutputEl = document.getElementById("singleOutput");
  const errorBoxEl = document.getElementById("errorBox");
  const dropZone = document.getElementById("dropZone");

  const onFilesSelected = () => {
    const hasFiles = Boolean(filesInput.files && filesInput.files.length > 0);
    convertBtn.disabled = !hasFiles;
    statusEl.textContent = hasFiles
      ? `${filesInput.files.length} ENEX file(s) selected.`
      : "";
    if (!hasFiles) {
      updateProgress(progressBarEl, progressTextEl, 0, 1, "No conversion running.");
      errorBoxEl.hidden = true;
      errorBoxEl.textContent = "";
    }
  };

  filesInput.addEventListener("change", onFilesSelected);
  initDragAndDrop(dropZone, filesInput, onFilesSelected);

  convertBtn.addEventListener("click", async () => {
    const files = Array.from(filesInput.files || []);
    if (!files.length) {
      statusEl.textContent = "Please choose at least one ENEX file.";
      return;
    }

    convertBtn.disabled = true;
    updateProgress(progressBarEl, progressTextEl, 0, 1, "Preparing conversion...");
    try {
      await convertFiles(files, {
        statusEl,
        progressBarEl,
        progressTextEl,
        singleOutputEnabled: singleOutputEl.checked,
        errorBoxEl
      });
    } catch (err) {
      statusEl.textContent = `Error: ${err.message || String(err)}`;
      errorBoxEl.hidden = false;
      errorBoxEl.textContent = `Conversion stopped: ${err.message || String(err)}`;
    } finally {
      convertBtn.disabled = false;
    }
  });

  updateProgress(progressBarEl, progressTextEl, 0, 1, "No conversion running.");
}

window.addEventListener("DOMContentLoaded", init);
