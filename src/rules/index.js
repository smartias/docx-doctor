import { numberingRestartRule } from "./numbering-restart.js";
import { headingPPrOrderRule } from "./heading-ppr-order.js";
import { trailingBlankPagesRule } from "./trailing-blank-pages.js";
import { splitRunRiskRule } from "./split-run-risk.js";

/** Built-in rule registry, in the order they were built (see ROADMAP.md). */
export const RULES = [
  trailingBlankPagesRule,
  headingPPrOrderRule,
  numberingRestartRule,
  splitRunRiskRule,
];
