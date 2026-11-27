import { createSheetEditor, createSheetEditorAsync, setLocale, type SheetEditor } from "../src/index";

// ?lang=fr opens the demo in another language, with its numbers and formula names.
const lang = new URLSearchParams(location.search).get("lang");
if (lang) setLocale(lang);

const fileInput = document.getElementById("file") as HTMLInputElement;
const openBtn = document.getElementById("open") as HTMLButtonElement;
const saveBtn = document.getElementById("save") as HTMLButtonElement;
const statusEl = document.getElementById("status") as HTMLElement;
const editorEl = document.getElementById("editor") as HTMLElement;

openBtn.addEventListener("click", () => fileInput.click());

let editor: SheetEditor | null = null;
let filename = "edited.xlsx";

async function loadBytes(bytes: Uint8Array, name: string): Promise<void> {
  filename = name;
  editor?.destroy();
  editorEl.innerHTML = "";
  try {
    const win = window as unknown as Record<string, unknown>;
    win.seChangeCount = 0; // lets a host (and the e2e suite) see that onChange really fired
    win.seCellChanges = [];
    win.seStructural = [];
    // The async variant asks for the password of an encrypted .xls with its own prompt.
    editor = await createSheetEditorAsync(editorEl, bytes, {
      fonts: { remote: { provider: "google" } }, // the demo offers "Get more fonts"
      onChange: () => {
        win.seChangeCount = (win.seChangeCount as number) + 1;
        saveBtn.disabled = false;
        statusEl.textContent = "edited";
      },
      onCellsChanged: (changes) => {
        (win.seCellChanges as unknown[]).push(...changes);
      },
      onSelectionChanged: (at) => {
        win.seSelection = at;
      },
      allowStructuralEdit: (op) => {
        (win.seStructural as unknown[]).push(op);
        return win.seBlockStructural !== true;
      },
    });
    win.seHandle = editor; // handy in the console, and how a host would drive it
    // Lets a test build a second editor on the page and wire the two together, which is
    // how a collaboration host uses this library. Doing that in the browser is the only
    // way to check the pair without a network in the way.
    win.createSheetEditor = createSheetEditor;
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
  // A legacy .xls may hold things BIFF8 saving cannot keep; say so before writing it.
  const warnings = editor.saveWarnings();
  if (warnings.length && !confirm(warnings.join("\n\n"))) return;
  const out = await editor.getBytes();
  const lower = filename.toLowerCase();
  const ext = lower.endsWith(".ods")
    ? "application/vnd.oasis.opendocument.spreadsheet"
    : lower.endsWith(".xls") ? "application/vnd.ms-excel"
    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const blob = new Blob([out as BlobPart], { type: ext });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
});
