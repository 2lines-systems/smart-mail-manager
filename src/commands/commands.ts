import { loadConfiguration } from "../services/ConfigurationService";
import { getSignatureHtml } from "../services/SignatureService";
import { log, debug, error, setDebugLogging } from "../utils/logger";
import { Configuration } from "../models/Configuration";

// Globaler Initialisierungsstatus
let initialized = false;
let config: Configuration | null = null;

/**
 * Initialisierung: Config laden und Debug-Logging setzen.
 * Wird bei jedem Event-Handler-Aufruf aufgerufen (idempotent durch initialized-Flag).
 */
async function ensureInitialized(): Promise<void> {
  if (initialized && config) return;

  try {
    config = await loadConfiguration();
    setDebugLogging(config.enableDebugLogging);
    initialized = true;
    log("Initialisierung abgeschlossen");
  } catch (err) {
    error("Initialisierung fehlgeschlagen", err);
  }
}

/**
 * Liest die aktuelle Absender-Adresse aus dem Compose-Item.
 */
function getFromAddress(): Promise<string | null> {
  return new Promise((resolve) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        resolve(null);
        return;
      }

      item.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          resolve(result.value.emailAddress);
        } else {
          resolve(null);
        }
      });
    } catch {
      resolve(null);
    }
  });
}

/**
 * Setzt die Signatur im aktuellen Compose-Item.
 */
function setSignature(html: string): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.item!.body.setSignatureAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            error("setSignatureAsync fehlgeschlagen", result.error);
            resolve(false);
          }
        },
      );
    } catch (err) {
      error("Signatur setzen fehlgeschlagen", err);
      resolve(false);
    }
  });
}

/**
 * Kern-Logik: Absender lesen → Signatur matchen → Signatur setzen.
 * Wird von OnNewMessageCompose und OnMessageFromChanged verwendet.
 */
async function applySignatureForCurrentSender(): Promise<void> {
  await ensureInitialized();
  if (!config) return;

  const senderEmail = await getFromAddress();
  if (!senderEmail) {
    debug("Kein Absender ermittelt");
    return;
  }

  log(`Absender: ${senderEmail}`);

  const signatureHtml = await getSignatureHtml(senderEmail, config);
  if (!signatureHtml) {
    debug("Keine Signatur fuer diesen Absender");
    return;
  }

  const success = await setSignature(signatureHtml);
  if (success) {
    log(`Signatur gesetzt fuer ${senderEmail}`);
  }
}

// =============================================================================
// Event-Handler
// =============================================================================

/**
 * OnNewMessageCompose: Wird ausgeloest wenn eine neue Mail erstellt wird
 * (Neue Mail, Antwort, Weiterleitung).
 */
async function onNewMessageComposeHandler(event: Office.AddinCommands.Event): Promise<void> {
  try {
    log("Event: OnNewMessageCompose");
    await applySignatureForCurrentSender();
  } catch (err) {
    error("Fehler in OnNewMessageCompose", err);
  } finally {
    event.completed();
  }
}

/**
 * OnMessageFromChanged: Wird ausgeloest wenn der Benutzer das "Von"-Feld aendert.
 */
async function onMessageFromChangedHandler(event: Office.AddinCommands.Event): Promise<void> {
  try {
    log("Event: OnMessageFromChanged");
    await applySignatureForCurrentSender();
  } catch (err) {
    error("Fehler in OnMessageFromChanged", err);
  } finally {
    event.completed();
  }
}

/**
 * OnMessageSend (Smart Alert): Prueft auf externe Empfaenger vor dem Senden.
 */
async function onMessageSendHandler(event: Office.AddinCommands.Event): Promise<void> {
  try {
    log("Event: OnMessageSend");
    await ensureInitialized();

    if (!config?.showWarningForExternal) {
      event.completed({ allowEvent: true });
      return;
    }

    const hasExternal = await checkExternalRecipients();
    if (hasExternal) {
      const senderEmail = await getFromAddress();
      log(`Externe Empfaenger erkannt — Warnung anzeigen (Absender: ${senderEmail})`);

      event.completed({
        allowEvent: false,
        errorMessage:
          `Diese Nachricht enthaelt externe Empfaenger.\n\n` +
          `Absender: ${senderEmail || "(unbekannt)"}\n\n` +
          `Bitte pruefen Sie, ob der Absender korrekt ist.`,
      } as Office.AddinCommands.EventCompletedOptions);
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (err) {
    error("Fehler in OnMessageSend", err);
    // Im Fehlerfall Senden erlauben (Graceful Degradation)
    event.completed({ allowEvent: true });
  }
}

/**
 * Prueft ob externe Empfaenger vorhanden sind.
 */
async function checkExternalRecipients(): Promise<boolean> {
  if (!config?.internalDomains || config.internalDomains.length === 0) {
    return false;
  }

  const allRecipients = await getAllRecipients();
  const normalizedDomains = config.internalDomains.map((d) => d.toLowerCase());

  for (const recipient of allRecipients) {
    const email = recipient.toLowerCase();
    const isInternal = normalizedDomains.some((domain) => email.endsWith(domain));
    if (!isInternal) {
      debug(`Externer Empfaenger: ${recipient}`);
      return true;
    }
  }

  return false;
}

/**
 * Liest alle Empfaenger (To, CC, BCC) aus dem aktuellen Item.
 */
function getAllRecipients(): Promise<string[]> {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      resolve([]);
      return;
    }

    const emails: string[] = [];
    let pending = 3;

    const collect = (result: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        for (const r of result.value) {
          if (r.emailAddress) {
            emails.push(r.emailAddress);
          }
        }
      }
      pending--;
      if (pending === 0) {
        resolve(emails);
      }
    };

    item.to.getAsync(collect);
    item.cc.getAsync(collect);
    item.bcc.getAsync(collect);
  });
}

// =============================================================================
// Event-Handler registrieren
// =============================================================================

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
