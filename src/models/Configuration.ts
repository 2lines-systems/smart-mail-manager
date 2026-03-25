/**
 * Gesamtkonfiguration des Smart Mail Manager Web Add-ins.
 * Kompatibel mit der VSTO-Konfiguration, erweitert um URL-basierte Felder.
 */
export interface Configuration {
  /** Pfad zur zentralen Config (nur fuer Dokumentation/Kompatibilitaet mit VSTO) */
  centralConfigPath?: string;

  /** HTTPS-URL zur zentralen Config (primaer fuer Web Add-in) */
  centralConfigUrl?: string;

  /** Konfigurierte Postfaecher */
  mailboxes: MailboxConfig[];

  /** Interne Domains z.B. ["@firma.de", "@tochter.de"] */
  internalDomains: string[];

  /** Warnung bei externen Empfaengern anzeigen */
  showWarningForExternal: boolean;

  /** Signaturen automatisch anwenden */
  signatureManagementEnabled: boolean;

  /** Basis-URL zu Signatur-Dateien (z.B. SharePoint-URL) */
  signaturesBaseUrl?: string;

  /** Fallback-Signatur wenn kein Mailbox-Match */
  defaultSignatureFile: string;

  /** Signatur-Position bei Antworten: "BeforeQuote" oder "AfterQuote" */
  replySignaturePosition: "BeforeQuote" | "AfterQuote";

  /** Debug-Logging aktivieren */
  enableDebugLogging: boolean;

  /** Profildaten-Fallback fuer Felder die nicht in Graph API verfuegbar sind */
  userProfileOverrides?: Record<string, string>;
}

/**
 * Konfiguration einer einzelnen Shared Mailbox.
 * Gegenueber VSTO: emailAddress ist jetzt das primaere Match-Kriterium (statt folderName).
 */
export interface MailboxConfig {
  /** SMTP-Adresse der Mailbox (Match-Kriterium) */
  emailAddress: string;

  /** Aktiv? */
  enabled: boolean;

  /** HTML-Signatur-Dateiname (z.B. "info.htm") */
  signatureFile: string;
}

/** Standard-Konfiguration */
export const DEFAULT_CONFIGURATION: Configuration = {
  mailboxes: [],
  internalDomains: [],
  showWarningForExternal: false,
  signatureManagementEnabled: true,
  defaultSignatureFile: "default.htm",
  replySignaturePosition: "BeforeQuote",
  enableDebugLogging: false,
};
