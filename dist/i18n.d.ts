/** Force a locale (host escape hatch). Unknown codes fall back to English. */
export declare function setLocale(code: string): void;
export declare function t(key: string, params?: Record<string, string | number>): string;
