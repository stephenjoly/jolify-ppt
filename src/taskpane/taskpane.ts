import type { ActionResult } from "../shared/shapeTools";
import {
  toggleDraftSticker,
  alignBottom,
  copyFillToClipboard,
  pasteFillFromClipboard,
  matchFillStyle,
  copyOutlineToClipboard,
  pasteOutlineFromClipboard,
  matchOutlineStyle,
  copyStyleToClipboard,
  pasteStyleFromClipboard,
  matchShapeStyle,
  clearFill,
  clearOutline,
  matchHeight,
  matchHeightAndWidth,
  matchWidth,
  alignCenterH,
  alignLeft,
  alignMiddleV,
  alignRight,
  alignTop,
  copyPositionAndSize,
  copyPositionOnly,
  copySizeOnly,
  distributeH,
  distributeV,
  stretchToLeftEdge,
  stretchToRightEdge,
  stretchToTopEdge,
  stretchToBottomEdge,
  pastePositionAndSize,
  pastePositionOnly,
  pasteSizeOnly,
  splitTextBoxByLines,
  removeTextMargins,
  disableTextAutofit,
  createCenterSticker,
  swapPositions,
  // Phase 1 — new
  mergeTextBoxes,
  alignCenterHAndGroup,
  alignMiddleVAndGroup,
  centerMiddleAndGroup,
  distributeHandVAndGroup,
  distributeHAndGroup,
  distributeVAndGroup,
  // Phase 2 — dialog wrappers
  openGridDialog,
  openWeekdayRangeDialog,
  openSelectedDeckDialog,
  moveToUnusedSection,
} from "../shared/shapeTools";

type ActionRunner = () => Promise<ActionResult>;

const ACTIONS: Record<string, ActionRunner> = {
  "copy-all-btn": copyPositionAndSize,
  "paste-all-btn": pastePositionAndSize,
  "copy-position-btn": copyPositionOnly,
  "paste-position-btn": pastePositionOnly,
  "copy-size-btn": copySizeOnly,
  "paste-size-btn": pasteSizeOnly,
  "swap-btn": swapPositions,
  "copy-fill-btn":       copyFillToClipboard,
  "paste-fill-btn":      pasteFillFromClipboard,
  "match-fill-btn":      matchFillStyle,
  "copy-outline-btn":    copyOutlineToClipboard,
  "paste-outline-btn":   pasteOutlineFromClipboard,
  "match-outline-btn":   matchOutlineStyle,
  "copy-style-btn":      copyStyleToClipboard,
  "paste-style-btn":     pasteStyleFromClipboard,
  "match-style-btn":     matchShapeStyle,
  "clear-fill-btn":      clearFill,
  "clear-outline-btn":   clearOutline,
  "match-height-btn":    matchHeight,
  "match-width-btn":     matchWidth,
  "match-size-btn":      matchHeightAndWidth,
  "align-left-btn":   alignLeft,
  "align-center-btn": alignCenterH,
  "align-right-btn":  alignRight,
  "align-top-btn":    alignTop,
  "align-middle-btn": alignMiddleV,
  "align-bottom-btn": alignBottom,
  "distribute-h-btn": distributeH,
  "distribute-v-btn": distributeV,
  "stretch-left-btn": stretchToLeftEdge,
  "stretch-right-btn": stretchToRightEdge,
  "stretch-top-btn": stretchToTopEdge,
  "stretch-bottom-btn": stretchToBottomEdge,
  "toggle-draft-btn": toggleDraftSticker,
  // Align & Group
  "align-center-group-btn":  alignCenterHAndGroup,
  "align-middle-group-btn":  alignMiddleVAndGroup,
  "center-middle-group-btn": centerMiddleAndGroup,
  "distribute-hv-group-btn": distributeHandVAndGroup,
  "distribute-h-group-btn":  distributeHAndGroup,
  "distribute-v-group-btn":  distributeVAndGroup,
  // Text Tools
  "merge-textboxes-btn": mergeTextBoxes,
  "split-textbox-btn": splitTextBoxByLines,
  "remove-text-margins-btn": removeTextMargins,
  "disable-text-autofit-btn": disableTextAutofit,
  "create-center-sticker-btn": createCenterSticker,
  // Layout Builders
  "create-grid-btn":    openGridDialog,
  "weekday-range-btn": openWeekdayRangeDialog,
  // Slides
  "save-selected-deck-btn": openSelectedDeckDialog,
  "move-to-unused-btn": moveToUnusedSection,
};

function statusEl() {
  return document.getElementById("status")!;
}

function getAllButtons() {
  return Object.keys(ACTIONS)
    .map((id) => document.getElementById(id) as HTMLButtonElement | null)
    .filter((btn): btn is HTMLButtonElement => btn instanceof HTMLButtonElement);
}

function setStatus(result: ActionResult) {
  const el = statusEl();
  el.className = result.type;
  el.textContent = "";

  const labels: Record<ActionResult["type"], string> = {
    success: "Success",
    info: "Info",
    warning: "Heads up",
    error: "Error",
  };

  const label = document.createElement("strong");
  label.textContent = labels[result.type];
  el.append(label);

  const message = document.createElement("span");
  message.textContent = result.message;
  message.style.display = "block";
  el.append(message);
}

function setBusy(isBusy: boolean) {
  getAllButtons().forEach((btn) => {
    btn.disabled = isBusy;
  });

  if (isBusy) {
    const el = statusEl();
    el.className = "info";
    el.textContent = "Working...";
  }
}

function wireAction(buttonId: string, runner: ActionRunner) {
  const button = document.getElementById(buttonId) as HTMLButtonElement | null;
  if (!button) {
    return;
  }

  button.addEventListener("click", async () => {
    setBusy(true);
    try {
      const result = await runner();
      setStatus(result);
    } catch (error) {
      console.error(error);
      setStatus({
        type: "error",
        message: error instanceof Error ? error.message : "Something went wrong.",
      });
    } finally {
      setBusy(false);
    }
  });
}

function applyNativeTooltips() {
  document.querySelectorAll<HTMLButtonElement>(".tool-button").forEach((button) => {
    const title = button.querySelector(".tooltip-title")?.textContent?.trim() ?? "";
    const copy = button.querySelector(".tooltip-copy")?.textContent?.trim() ?? "";
    const tooltip = copy ? `${title}\n${copy}` : title;

    if (tooltip) {
      button.title = tooltip;
      button.setAttribute("aria-description", copy || title);
    }
  });
}

Office.onReady(() => {
  applyNativeTooltips();

  Object.entries(ACTIONS).forEach(([id, runner]) => {
    wireAction(id, runner);
  });

  setStatus({
    type: "info",
    message: "Ready when you are.",
  });
});
