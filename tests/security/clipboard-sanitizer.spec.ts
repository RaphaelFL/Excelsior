import { describe, expect, it } from "vitest";
import { resolveClipboardText } from "@excelsior/renderer-dom";

describe("clipboard security", () => {
  it("blocks unsafe html when policy requires blocking", () => {
    const analysis = resolveClipboardText(
      "Safe",
      "<div onclick='hack()'><script>alert(1)</script>Safe</div>",
      "blocked-html"
    );

    expect(analysis.blocked).toBe(true);
    expect(analysis.hardBlocked).toBe(true);
    expect(analysis.reasons.length).toBeGreaterThan(0);
  });

  it("blocks large payloads even without html", () => {
    const analysis = resolveClipboardText("A\tB\n1\t2", "", "safe-html", 3);

    expect(analysis.hardBlocked).toBe(true);
    expect(analysis.reasons).toContain("max-paste-cells-exceeded:4:3");
  });
});