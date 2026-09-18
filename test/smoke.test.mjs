import { test } from "node:test";
import assert from "node:assert/strict";
import { RULES } from "../src/rules/index.js";

// This is the one test that should pass on day zero — it just proves the
// project scaffold and rule registry load correctly. Replace/extend with
// real fixture-backed tests as each rule gets implemented (see ROADMAP.md).
test("rule registry loads and each rule has the expected shape", () => {
  assert.ok(RULES.length > 0);
  for (const rule of RULES) {
    assert.equal(typeof rule.id, "string");
    assert.ok(rule.severity === "error" || rule.severity === "warn");
    assert.equal(typeof rule.detect, "function");
    // repair() is optional — split-run-risk is detect-only by design.
    if (rule.repair !== undefined) {
      assert.equal(typeof rule.repair, "function");
    }
  }
});
