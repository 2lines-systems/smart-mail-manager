import { Configuration, DEFAULT_CONFIGURATION } from "../models/Configuration";
import { log, debug, error } from "../utils/logger";

const CACHE_KEY = "SmartMailManager_Config";
const CACHE_TIMESTAMP_KEY = "SmartMailManager_ConfigTimestamp";
const CACHE_MAX_AGE_MS = 60 * 60 * 1000; // 1 Stunde

let cachedConfig: Configuration | null = null;

/**
 * Laedt die Konfiguration.
 * Strategie: RoamingSettings-Cache → centralConfigUrl → Default
 */
export async function loadConfiguration(): Promise<Configuration> {
  if (cachedConfig) {
    debug("Config aus In-Memory-Cache verwendet");
    return cachedConfig;
  }

  // 1. RoamingSettings-Cache pruefen
  const cached = loadFromRoamingSettings();
  if (cached) {
    log("Config aus RoamingSettings-Cache geladen");
    cachedConfig = cached;
    return cached;
  }

  // 2. Zentrale Config per URL laden
  const bootstrapConfig = loadBootstrapConfig();
  if (bootstrapConfig?.centralConfigUrl) {
    try {
      const centralConfig = await fetchConfig(bootstrapConfig.centralConfigUrl);
      if (centralConfig) {
        const merged = mergeConfigs(centralConfig, bootstrapConfig);
        saveToRoamingSettings(merged);
        cachedConfig = merged;
        log("Zentrale Config geladen und gecached");
        return merged;
      }
    } catch (err) {
      error("Zentrale Config konnte nicht geladen werden", err);
    }
  }

  // 3. Fallback: Default-Config
  log("Verwende Default-Konfiguration");
  cachedConfig = { ...DEFAULT_CONFIGURATION };
  return cachedConfig;
}

/**
 * Gibt die aktuelle Config zurueck (ohne erneutes Laden).
 */
export function getConfiguration(): Configuration | null {
  return cachedConfig;
}

/**
 * Erzwingt ein Neuladen der Config beim naechsten Aufruf.
 */
export function invalidateCache(): void {
  cachedConfig = null;
  try {
    Office.context.roamingSettings.remove(CACHE_KEY);
    Office.context.roamingSettings.remove(CACHE_TIMESTAMP_KEY);
    Office.context.roamingSettings.saveAsync();
    debug("Config-Cache invalidiert");
  } catch {
    // RoamingSettings nicht verfuegbar — ignorieren
  }
}

/**
 * Laedt Config aus RoamingSettings, wenn nicht aelter als CACHE_MAX_AGE_MS.
 */
function loadFromRoamingSettings(): Configuration | null {
  try {
    const timestamp = Office.context.roamingSettings.get(CACHE_TIMESTAMP_KEY) as number | undefined;
    if (!timestamp || Date.now() - timestamp > CACHE_MAX_AGE_MS) {
      return null;
    }
    const data = Office.context.roamingSettings.get(CACHE_KEY) as string | undefined;
    if (!data) return null;
    return JSON.parse(data) as Configuration;
  } catch {
    return null;
  }
}

/**
 * Speichert Config in RoamingSettings mit Timestamp.
 */
function saveToRoamingSettings(config: Configuration): void {
  try {
    Office.context.roamingSettings.set(CACHE_KEY, JSON.stringify(config));
    Office.context.roamingSettings.set(CACHE_TIMESTAMP_KEY, Date.now());
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        error("RoamingSettings speichern fehlgeschlagen", result.error);
      }
    });
  } catch {
    // RoamingSettings nicht verfuegbar — ignorieren
  }
}

/**
 * Laedt Bootstrap-Config aus RoamingSettings.
 * Die Bootstrap-Config enthaelt mindestens die centralConfigUrl
 * und wird beim initialen Deployment gesetzt.
 */
function loadBootstrapConfig(): Partial<Configuration> | null {
  try {
    const data = Office.context.roamingSettings.get("SmartMailManager_Bootstrap") as string | undefined;
    if (!data) return null;
    return JSON.parse(data) as Partial<Configuration>;
  } catch {
    return null;
  }
}

/**
 * Laedt Config per HTTPS-URL (SharePoint, Azure Blob, etc.).
 */
async function fetchConfig(url: string): Promise<Configuration | null> {
  debug(`Lade Config von: ${url}`);

  const response = await fetch(url, {
    headers: { Accept: "application/json" },
    cache: "no-cache",
  });

  if (!response.ok) {
    error(`Config-Fetch fehlgeschlagen: ${response.status} ${response.statusText}`);
    return null;
  }

  const data = await response.json();
  return data as Configuration;
}

/**
 * Merged zentrale Config mit Bootstrap-Overrides.
 * Gleiche Logik wie VSTO MergeLocalOverrides: Nur explizit gesetzte Felder ueberschreiben.
 */
function mergeConfigs(central: Configuration, local: Partial<Configuration>): Configuration {
  const merged = { ...central };

  if (local.enableDebugLogging !== undefined) merged.enableDebugLogging = local.enableDebugLogging;
  if (local.showWarningForExternal !== undefined) merged.showWarningForExternal = local.showWarningForExternal;
  if (local.signatureManagementEnabled !== undefined) merged.signatureManagementEnabled = local.signatureManagementEnabled;
  if (local.internalDomains !== undefined && local.internalDomains.length > 0) merged.internalDomains = local.internalDomains;
  if (local.signaturesBaseUrl !== undefined) merged.signaturesBaseUrl = local.signaturesBaseUrl;
  if (local.defaultSignatureFile !== undefined) merged.defaultSignatureFile = local.defaultSignatureFile;
  if (local.mailboxes !== undefined && local.mailboxes.length > 0) merged.mailboxes = local.mailboxes;
  if (local.userProfileOverrides !== undefined) merged.userProfileOverrides = local.userProfileOverrides;

  // centralConfigUrl immer aus Bootstrap beibehalten
  if (local.centralConfigUrl) merged.centralConfigUrl = local.centralConfigUrl;

  return merged;
}
