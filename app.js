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

async function convertFiles(files, statusEl) {
  const zip = new JSZip();
  let totalNotes = 0;

  for (const file of files) {
    statusEl.textContent = `Reading ${file.name}...`;
    const xmlText = await readFileAsText(file);
    const notes = parseEnex(xmlText);

    const notebookFolder = zip.folder(cleanFileName(file.name.replace(/\.enex$/i, "")));
    let noteIndex = 1;

    for (const note of notes) {
      totalNotes += 1;
      statusEl.textContent = `Converting: ${note.title}`;
      const html = buildDocHtml(note);
      const blob = window.htmlDocx.asBlob(html);
      const fileName = `${String(noteIndex).padStart(4, "0")}-${cleanFileName(note.title)}.docx`;
      notebookFolder.file(fileName, blob);
      noteIndex += 1;
    }
  }

  statusEl.textContent = "Creating ZIP...";
  const finalZip = await zip.generateAsync({ type: "blob" });
  saveAs(finalZip, `enex-to-onenote-${Date.now()}.zip`);
  statusEl.textContent = `Done. Converted ${totalNotes} notes.`;
}

function init() {
  const filesInput = document.getElementById("enexFiles");
  const convertBtn = document.getElementById("convertBtn");
  const statusEl = document.getElementById("status");

  filesInput.addEventListener("change", () => {
    convertBtn.disabled = !filesInput.files || filesInput.files.length === 0;
    statusEl.textContent = filesInput.files?.length
      ? `${filesInput.files.length} file(s) selected.`
      : "";
  });

  convertBtn.addEventListener("click", async () => {
    const files = Array.from(filesInput.files || []);
    if (!files.length) {
      statusEl.textContent = "Please choose at least one ENEX file.";
      return;
    }

    convertBtn.disabled = true;
    try {
      await convertFiles(files, statusEl);
    } catch (err) {
      statusEl.textContent = `Error: ${err.message || String(err)}`;
    } finally {
      convertBtn.disabled = false;
    }
  });
}

window.addEventListener("DOMContentLoaded", init);
