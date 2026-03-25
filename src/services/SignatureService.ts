import { Configuration, MailboxConfig } from "../models/Configuration";
import { getUserProfile, getPlaceholderValue } from "./GraphService";
import { log, debug, error } from "../utils/logger";

/** Cache fuer geladene Signatur-HTML-Dateien */
const signatureCache = new Map<string, string>();

/**
 * Ermittelt und bereitet die Signatur-HTML fuer eine Absenderadresse vor.
 * 1. Mailbox-Match → signatureFile
 * 2. Fallback: defaultSignatureFile
 * 3. HTML laden (Cache oder Fetch)
 * 4. Platzhalter ersetzen
 */
export async function getSignatureHtml(
  senderEmail: string,
  config: Configuration,
): Promise<string | null> {
  if (!config.signatureManagementEnabled) {
    debug("Signaturverwaltung deaktiviert");
    return null;
  }

  // Signatur-Dateiname bestimmen
  const signatureFile = getSignatureFile(senderEmail, config);
  if (!signatureFile) {
    debug(`Keine Signatur-Datei fuer ${senderEmail}`);
    return null;
  }

  // HTML laden
  const html = await loadSignature(signatureFile, config);
  if (!html) {
    return null;
  }

  // Platzhalter ersetzen
  const replaced = await replacePlaceholders(html, config.userProfileOverrides);
  return replaced;
}

/**
 * Bestimmt den Signatur-Dateinamen fuer eine Absenderadresse.
 */
function getSignatureFile(senderEmail: string, config: Configuration): string | null {
  const normalizedSender = senderEmail.toLowerCase();

  // Match gegen konfigurierte Mailboxes
  const match = config.mailboxes.find(
    (mb: MailboxConfig) => mb.enabled && mb.signatureFile && mb.emailAddress.toLowerCase() === normalizedSender,
  );

  if (match) {
    debug(`Signatur-Match: ${senderEmail} → ${match.signatureFile}`);
    return match.signatureFile;
  }

  // Fallback
  if (config.defaultSignatureFile) {
    debug(`Kein Mailbox-Match fuer ${senderEmail}, verwende Default: ${config.defaultSignatureFile}`);
    return config.defaultSignatureFile;
  }

  return null;
}

/**
 * Laedt eine Signatur-HTML-Datei per URL (mit Cache).
 */
async function loadSignature(fileName: string, config: Configuration): Promise<string | null> {
  // Cache pruefen
  const cached = signatureCache.get(fileName);
  if (cached) {
    debug(`Signatur aus Cache: ${fileName}`);
    return cached;
  }

  if (!config.signaturesBaseUrl) {
    error("Kein signaturesBaseUrl konfiguriert — Signaturen koennen nicht geladen werden");
    return null;
  }

  const url = `${config.signaturesBaseUrl.replace(/\/+$/, "")}/${fileName}`;
  debug(`Lade Signatur: ${url}`);

  try {
    const response = await fetch(url, { cache: "no-cache" });
    if (!response.ok) {
      error(`Signatur laden fehlgeschlagen: ${response.status} ${url}`);
      return null;
    }

    const html = await response.text();
    signatureCache.set(fileName, html);
    log(`Signatur geladen: ${fileName} (${html.length} Zeichen)`);
    return html;
  } catch (err) {
    error(`Signatur konnte nicht geladen werden: ${fileName}`, err);
    return null;
  }
}

/**
 * Ersetzt {{Platzhalter}} in der Signatur-HTML.
 * Fallback-Kette: Graph API → Config-Overrides → Platzhalter entfernen
 */
async function replacePlaceholders(
  html: string,
  overrides?: Record<string, string>,
): Promise<string> {
  const profile = await getUserProfile();

  // Alle {{...}} Platzhalter finden und ersetzen
  const replaced = html.replace(/\{\{([^}]+)\}\}/g, (_match, name: string) => {
    const value = getPlaceholderValue(name.trim(), profile, overrides);
    if (value) {
      debug(`Platzhalter {{${name}}} → "${value}"`);
    } else {
      debug(`Platzhalter {{${name}}} → (entfernt)`);
    }
    return value;
  });

  return replaced;
}

/** Signatur-Cache leeren (z.B. bei Config-Reload) */
export function invalidateSignatureCache(): void {
  signatureCache.clear();
  debug("Signatur-Cache geleert");
}
