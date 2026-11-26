import { afterEach, describe, expect, it } from "vitest";
import { localeCode, localeDicts, setLocale, t } from "./i18n";

afterEach(() => setLocale("en"));

const placeholders = (s: string): string[] => (s.match(/\{\w+\}/g) ?? []).sort();

describe("UI dictionaries", () => {
  const dicts = localeDicts();
  const en = dicts.en!;

  for (const code of ["it", "pl", "ko", "tr"]) {
    it(`${code} has every English key and its placeholders`, () => {
      const d = dicts[code]!;
      expect(Object.keys(d)).toEqual(Object.keys(en));
      for (const [k, v] of Object.entries(en)) {
        expect(d[k], `${code}.${k}`).toBeTruthy();
        expect(placeholders(d[k]!), `${code}.${k}`).toEqual(placeholders(v));
      }
    });
  }

  it("never uses an em dash", () => {
    for (const code of ["it", "pl", "ko", "tr"]) {
      for (const v of Object.values(dicts[code]!)) expect(v).not.toContain("\u2014");
    }
  });

  it("switches to the new languages", () => {
    setLocale("it-IT");
    expect(localeCode()).toBe("it");
    expect(t("freezePanes")).toBe("Blocca riquadri");
    setLocale("pl");
    expect(t("nameManager")).toBe("Menedżer nazw");
    setLocale("ko-KR");
    expect(t("cfEdit")).toBe("조건부 서식");
    setLocale("tr");
    expect(t("dvEdit")).toBe("Veri doğrulama");
    expect(t("rowsDel", { n: 3 })).toBe("3 satırı sil");
  });
});
