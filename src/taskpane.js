// Task pane logic: shows the subject of the selected message and saves a new one via Microsoft Graph.
import { loadConfig } from "./config.js";
import { initAuth, getAccessToken } from "./auth.js";
import { getSubject, setSubject, GraphError } from "./graph.js";
import { setLanguage, t } from "./i18n.js";

// Exchange limits the subject to 255 characters.
const MAX_LENGTH = 255;

const el = {};
let currentRef = null;
let originalSubject = "";
let busy = false;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;

  setLanguage(Office.context.displayLanguage);
  bindElements();
  applyStrings();

  try {
    const config = await loadConfig();
    await initAuth(config);
  } catch (error) {
    showStatus(error.message === "config-missing" ? t("notConfigured") : t("signInFailed", { msg: error.message }), "error");
    el.form.hidden = true;
    return;
  }

  // Pinned task pane: react when the user selects another message (Mailbox 1.5).
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => loadItem());
  await loadItem();
});

function bindElements() {
  for (const id of ["title", "form", "currentLabel", "current", "newLabel", "subject", "counter", "save", "reset", "shortcut", "status"]) {
    el[id] = document.getElementById(id);
  }
  el.form.addEventListener("submit", (event) => {
    event.preventDefault();
    save();
  });
  el.reset.addEventListener("click", () => {
    el.subject.value = originalSubject;
    updateState();
    el.subject.focus();
  });
  el.subject.addEventListener("input", updateState);
  el.subject.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      save();
    } else if (event.key === "Enter") {
      // A subject is a single line; Enter must not insert a line break.
      event.preventDefault();
    }
  });
}

function applyStrings() {
  el.title.textContent = t("title");
  el.currentLabel.textContent = t("currentLabel");
  el.newLabel.textContent = t("newLabel");
  el.subject.placeholder = t("placeholder");
  el.save.textContent = t("save");
  el.reset.textContent = t("reset");
  el.shortcut.textContent = t("shortcut");
}

async function loadItem() {
  const item = Office.context.mailbox.item;
  clearStatus();

  if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
    currentRef = null;
    el.form.hidden = true;
    showStatus(t("noItem"), "info");
    return;
  }

  el.form.hidden = false;
  currentRef = {
    restId: Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0),
    mailbox: await getSharedMailbox(item),
  };

  // Office.js already knows the subject; show it immediately and avoid a Graph round trip.
  originalSubject = item.subject || "";
  el.current.textContent = originalSubject || "—";
  el.subject.value = originalSubject;
  updateState();
  el.subject.focus();
  el.subject.setSelectionRange(el.subject.value.length, el.subject.value.length);
}

// Returns the SMTP address of a shared or delegated mailbox, or null for the user's own mailbox.
function getSharedMailbox(item) {
  if (typeof item.getSharedPropertiesAsync !== "function") return Promise.resolve(null);
  return new Promise((resolve) => {
    item.getSharedPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value?.targetMailbox) {
        resolve(result.value.targetMailbox);
      } else {
        resolve(null);
      }
    });
  });
}

function validate(value) {
  if (!value.trim()) return t("empty");
  if (value.length > MAX_LENGTH) return t("tooLong", { max: MAX_LENGTH });
  return null;
}

function updateState() {
  const value = el.subject.value;
  el.counter.textContent = t("chars", { n: value.length, max: MAX_LENGTH });
  el.counter.classList.toggle("over", value.length > MAX_LENGTH);
  const changed = value !== originalSubject;
  el.save.disabled = busy || !changed || Boolean(validate(value));
  el.reset.disabled = busy || !changed;
}

async function save() {
  if (busy || !currentRef) return;
  const newSubject = el.subject.value.replace(/[\r\n]+/g, " ");

  const problem = validate(newSubject);
  if (problem) {
    showStatus(problem, "error");
    return;
  }
  if (newSubject === originalSubject) {
    showStatus(t("unchanged"), "info");
    return;
  }

  setBusy(true);
  showStatus(t("saving"), "info");
  const ref = currentRef;

  try {
    const token = await getAccessToken();
    await setSubject(token, ref, newSubject);
    // Read back what Exchange actually stored.
    const stored = await getSubject(token, ref);
    if (ref !== currentRef) return; // user switched to another message meanwhile

    originalSubject = stored;
    el.current.textContent = stored || "—";
    el.subject.value = stored;
    showStatus(t("saved"), "success");
  } catch (error) {
    showStatus(describeError(error), "error");
  } finally {
    setBusy(false);
  }
}

function describeError(error) {
  if (error instanceof GraphError) {
    if (error.status === 403) return t("forbidden");
    if (error.status === 404) return t("notFound");
    return t("graphError", { status: error.status, msg: error.message });
  }
  if (error.message === "dialog-blocked") return t("dialogBlocked");
  return t("signInFailed", { msg: error.message });
}

function setBusy(value) {
  busy = value;
  el.form.setAttribute("aria-busy", String(value));
  updateState();
}

function showStatus(text, kind) {
  el.status.textContent = text;
  el.status.dataset.kind = kind;
  el.status.hidden = false;
}

function clearStatus() {
  el.status.hidden = true;
  el.status.textContent = "";
}
