import { debug, error } from "../utils/logger";

/** Profildaten-Felder die von Graph API geladen werden */
export interface UserProfile {
  givenName?: string;
  surname?: string;
  displayName?: string;
  mail?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  department?: string;
  jobTitle?: string;
  companyName?: string;
  streetAddress?: string;
  postalCode?: string;
  city?: string;
}

/** Mapping von Platzhalter-Name auf Graph-Feld */
const PLACEHOLDER_MAP: Record<string, (p: UserProfile) => string | undefined> = {
  Vorname: (p) => p.givenName,
  Nachname: (p) => p.surname,
  Anzeigename: (p) => p.displayName,
  Email: (p) => p.mail,
  Telefon: (p) => p.businessPhones?.[0],
  Mobil: (p) => p.mobilePhone,
  Abteilung: (p) => p.department,
  Position: (p) => p.jobTitle,
  Firma: (p) => p.companyName,
  Strasse: (p) => p.streetAddress,
  PLZ: (p) => p.postalCode,
  Ort: (p) => p.city,
};

let cachedProfile: UserProfile | null = null;

/**
 * Laedt Benutzerprofil per Graph API.
 * Ergebnis wird pro Session gecacht (max 1 Aufruf).
 */
export async function getUserProfile(): Promise<UserProfile | null> {
  if (cachedProfile) {
    debug("Profil aus Cache verwendet");
    return cachedProfile;
  }

  try {
    const token = await getGraphToken();
    if (!token) {
      error("Kein Graph-Token verfuegbar");
      return null;
    }

    const select = [
      "givenName", "surname", "displayName", "mail",
      "businessPhones", "mobilePhone", "department",
      "jobTitle", "companyName", "streetAddress", "postalCode", "city",
    ].join(",");

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me?$select=${select}`,
      {
        headers: { Authorization: `Bearer ${token}` },
      },
    );

    if (!response.ok) {
      error(`Graph API Fehler: ${response.status} ${response.statusText}`);
      return null;
    }

    cachedProfile = (await response.json()) as UserProfile;
    debug("Profil per Graph API geladen");
    return cachedProfile;
  } catch (err) {
    error("Profil konnte nicht geladen werden", err);
    return null;
  }
}

/**
 * Gibt den Wert eines Platzhalters aus dem Profil zurueck.
 * Fallback auf overrides wenn Graph-Wert nicht verfuegbar.
 */
export function getPlaceholderValue(
  name: string,
  profile: UserProfile | null,
  overrides?: Record<string, string>,
): string {
  // 1. Graph-Profil
  if (profile) {
    const getter = PLACEHOLDER_MAP[name];
    if (getter) {
      const value = getter(profile);
      if (value) return value;
    }
  }

  // 2. Config-Override
  if (overrides?.[name]) {
    return overrides[name];
  }

  // 3. Kein Wert — Platzhalter wird entfernt
  return "";
}

/**
 * Holt Graph API Access Token ueber Office SSO.
 */
async function getGraphToken(): Promise<string | null> {
  try {
    // Office SSO: Holt ein ID-Token das gegen Graph getauscht werden kann
    const bootstrapToken = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
    });

    // In einer Produktions-Umgebung muss dieses Bootstrap-Token
    // serverseitig gegen ein Graph-Token getauscht werden (OBO-Flow).
    // Fuer Sideloading/Dev kann das Token direkt verwendet werden
    // wenn die AAD App entsprechend konfiguriert ist.
    return bootstrapToken;
  } catch (err) {
    const authError = err as { code?: string };
    if (authError.code === "13003") {
      debug("SSO nicht unterstuetzt — Graph-Features deaktiviert");
    } else {
      error("SSO Token-Fehler", err);
    }
    return null;
  }
}

/** Cache invalidieren (z.B. bei Session-Wechsel) */
export function invalidateProfileCache(): void {
  cachedProfile = null;
}
