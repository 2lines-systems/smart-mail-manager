let debugEnabled = false;

export function setDebugLogging(enabled: boolean): void {
  debugEnabled = enabled;
}

export function log(message: string): void {
  const timestamp = new Date().toISOString().substring(11, 23);
  const formatted = `[${timestamp}] [SmartMailManager] ${message}`;
  console.log(formatted);
}

export function debug(message: string): void {
  if (debugEnabled) {
    log(message);
  }
}

export function error(message: string, err?: unknown): void {
  const timestamp = new Date().toISOString().substring(11, 23);
  const formatted = `[${timestamp}] [SmartMailManager] ERROR: ${message}`;
  if (err) {
    console.error(formatted, err);
  } else {
    console.error(formatted);
  }
}
