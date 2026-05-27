import { describe, expect, it } from "vitest";
import { parseTabularText, resolveClipboardText, sanitizeClipboardHtml } from "../src/index";

describe("clipboard sanitizer", () => {
  it("extracts plain text from HTML tables", () => {
    const analysis = sanitizeClipboardHtml("<table><tr><td>A</td><td>B</td></tr></table>");

    expect(analysis.text).toBe("A\tB");
    expect(analysis.blocked).toBe(false);
  });

  it("flags blocked content from unsafe html", () => {
    const analysis = sanitizeClipboardHtml("<div onclick='x()'><script>alert(1)</script>Safe</div>");

    expect(analysis.blocked).toBe(true);
    expect(analysis.hardBlocked).toBe(false);
    expect(analysis.reasons).toContain("blocked-tag:script");
    expect(analysis.reasons).toContain("blocked-attribute:onclick");
  });

  it("flags disallowed tags and attributes outside the clipboard allowlist", () => {
    const analysis = sanitizeClipboardHtml("<table><tr><td style='color:red'><img src='x' />A</td></tr></table>");

    expect(analysis.blocked).toBe(true);
    expect(analysis.reasons).toContain("blocked-tag:img");
    expect(analysis.reasons).toContain("blocked-attribute:style");
    expect(analysis.reasons).toContain("blocked-attribute:src");
  });

  it("blocks oversized clipboard payloads", () => {
    const analysis = resolveClipboardText("A\tB\n1\t2", "", "text-only", 3);

    expect(analysis.blocked).toBe(true);
    expect(analysis.hardBlocked).toBe(true);
    expect(analysis.reasons).toContain("max-paste-cells-exceeded:4:3");
  });

  it("parses TSV matrices", () => {
    expect(parseTabularText("A\tB\n1\t2")).toEqual([
      ["A", "B"],
      ["1", "2"]
    ]);
  });
});