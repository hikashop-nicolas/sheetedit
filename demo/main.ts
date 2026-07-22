import { createSheetEditor, type SheetEditor } from "../src/index";

const fileInput = document.getElementById("file") as HTMLInputElement;
const openBtn = document.getElementById("open") as HTMLButtonElement;
const saveBtn = document.getElementById("save") as HTMLButtonElement;
const statusEl = document.getElementById("status") as HTMLElement;
const editorEl = document.getElementById("editor") as HTMLElement;

openBtn.addEventListener("click", () => fileInput.click());

let editor: SheetEditor | null = null;
let filename = "edited.xlsx";

function loadBytes(bytes: Uint8Array, name: string): void {
  filename = name;
  editor?.destroy();
  editorEl.innerHTML = "";
  try {
    editor = createSheetEditor(editorEl, bytes, {
      onChange: () => {
        saveBtn.disabled = false;
        statusEl.textContent = "edited";
      },
    });
    saveBtn.disabled = false;
    statusEl.textContent = "loaded";
  } catch (e) {
    statusEl.textContent = "could not open: " + (e as Error).message;
  }
}

fileInput.addEventListener("change", async () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  loadBytes(new Uint8Array(await file.arrayBuffer()), file.name);
});

// ?src=<url> loads a spreadsheet on startup (handy for quick tests).
const src = new URLSearchParams(location.search).get("src");
if (src) {
  fetch(src)
    .then(async (r) => loadBytes(new Uint8Array(await r.arrayBuffer()), src.split("/").pop() ?? "remote.xlsx"))
    .catch((e) => (statusEl.textContent = "could not fetch: " + (e as Error).message));
}

saveBtn.addEventListener("click", async () => {
  if (!editor) return;
  const out = await editor.getBytes();
  const ext = filename.toLowerCase().endsWith(".ods")
    ? "application/vnd.oasis.opendocument.spreadsheet"
    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const blob = new Blob([out as BlobPart], { type: ext });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
});
