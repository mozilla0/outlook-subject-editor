// UI strings. German (Swiss spelling) is the default, English is used for all other display languages.
const strings = {
  de: {
    title: "Betreff ändern",
    currentLabel: "Aktueller Betreff",
    newLabel: "Neuer Betreff",
    placeholder: "Neuen Betreff eingeben",
    save: "Betreff speichern",
    reset: "Zurücksetzen",
    shortcut: "Strg+Enter speichert",
    loading: "Betreff wird geladen …",
    saving: "Betreff wird gespeichert …",
    saved: "Betreff gespeichert. Outlook zeigt ihn nach dem Neuladen der Nachricht an.",
    unchanged: "Der Betreff ist unverändert.",
    empty: "Der Betreff darf nicht leer sein.",
    tooLong: "Der Betreff darf höchstens {max} Zeichen lang sein.",
    noItem: "Wähle eine empfangene Nachricht aus, um ihren Betreff zu ändern.",
    notConfigured: "Das Add-In ist noch nicht eingerichtet: config.json fehlt oder enthält keine Client-ID. Wende dich an die IT.",
    signInFailed: "Anmeldung fehlgeschlagen: {msg}",
    dialogBlocked: "Das Anmeldefenster konnte nicht geöffnet werden. Erlaube Pop-ups für Outlook und versuche es erneut.",
    graphError: "Microsoft Graph hat die Änderung abgelehnt ({status}): {msg}",
    forbidden: "Keine Berechtigung, diese Nachricht zu ändern. Bei freigegebenen Postfächern braucht es Vollzugriff.",
    notFound: "Die Nachricht wurde nicht gefunden. Möglicherweise wurde sie verschoben oder gelöscht.",
    chars: "{n} / {max}",
  },
  en: {
    title: "Change subject",
    currentLabel: "Current subject",
    newLabel: "New subject",
    placeholder: "Enter a new subject",
    save: "Save subject",
    reset: "Reset",
    shortcut: "Ctrl+Enter saves",
    loading: "Loading subject …",
    saving: "Saving subject …",
    saved: "Subject saved. Outlook shows it after the message is reloaded.",
    unchanged: "The subject has not changed.",
    empty: "The subject cannot be empty.",
    tooLong: "The subject can be at most {max} characters long.",
    noItem: "Select a received message to change its subject.",
    notConfigured: "The add-in is not set up yet: config.json is missing or has no client ID. Contact your IT team.",
    signInFailed: "Sign-in failed: {msg}",
    dialogBlocked: "The sign-in window could not be opened. Allow pop-ups for Outlook and try again.",
    graphError: "Microsoft Graph rejected the change ({status}): {msg}",
    forbidden: "You are not allowed to change this message. Shared mailboxes require full access.",
    notFound: "The message was not found. It may have been moved or deleted.",
    chars: "{n} / {max}",
  },
};

let lang = "de";

export function setLanguage(displayLanguage) {
  lang = (displayLanguage || "de").toLowerCase().startsWith("de") ? "de" : "en";
  document.documentElement.lang = lang === "de" ? "de-CH" : "en";
}

export function t(key, vars = {}) {
  const text = strings[lang][key] ?? strings.en[key] ?? key;
  return text.replace(/\{(\w+)\}/g, (_, name) => String(vars[name] ?? ""));
}
