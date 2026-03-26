import type { ActionResult } from "../shared/shapeTools";
import { getCurrentPresentationThemePalette } from "../shared/presentationTools";
import {
  toggleDraftSticker,
  alignBottom,
  applyFillColor,
  applyFontColor,
  applyOutlineColor,
  applyDefaultTextboxFormat,
  bringShapesForward,
  bringShapesToFront,
  cycleOutlineDashStyle,
  cycleOutlineWeight,
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
  distributeResizeHorizontal,
  distributeResizeVertical,
  distributeH,
  distributeV,
  groupSelectedShapes,
  insertArrow,
  insertElbowLine,
  insertOval,
  insertRectangle,
  stretchToLeftEdge,
  stretchToRightEdge,
  stretchToTopEdge,
  stretchToBottomEdge,
  insertStraightLine,
  insertTextBox,
  makeLinesHorizontal,
  makeLinesVertical,
  pastePositionAndSize,
  pastePositionOnly,
  pasteSizeOnly,
  sendShapesBackward,
  sendShapesToBack,
  setTextAutofitOff,
  setTextAutofitShapeToFitText,
  setTextAutofitTextToFitShape,
  setTextMarginsNone,
  setTextMarginsRoomy,
  setTextMarginsTight,
  splitTextBoxByLines,
  stackHorizontal,
  stackVertical,
  swapTextOnly,
  createCenterSticker,
  insertRoundedRectangle,
  swapXPositions,
  swapYPositions,
  swapPositions,
  toggleWordWrap,
  ungroupSelectedShapes,
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
type PaletteTarget = "font" | "outline" | "fill";

const ACTIONS: Record<string, ActionRunner> = {
  "copy-all-btn": copyPositionAndSize,
  "paste-all-btn": pastePositionAndSize,
  "copy-position-btn": copyPositionOnly,
  "paste-position-btn": pastePositionOnly,
  "copy-size-btn": copySizeOnly,
  "paste-size-btn": pasteSizeOnly,
  "swap-btn": swapPositions,
  "swap-x-btn": swapXPositions,
  "swap-y-btn": swapYPositions,
  "copy-fill-btn":       copyFillToClipboard,
  "paste-fill-btn":      pasteFillFromClipboard,
  "match-fill-btn":      matchFillStyle,
  "copy-outline-btn":    copyOutlineToClipboard,
  "paste-outline-btn":   pasteOutlineFromClipboard,
  "match-outline-btn":   matchOutlineStyle,
  "copy-style-btn":      copyStyleToClipboard,
  "paste-style-btn":     pasteStyleFromClipboard,
  "match-style-btn":     matchShapeStyle,
  "outline-weight-btn":  cycleOutlineWeight,
  "outline-dash-btn":    cycleOutlineDashStyle,
  "line-horizontal-btn": makeLinesHorizontal,
  "line-vertical-btn":   makeLinesVertical,
  "clear-fill-btn":      clearFill,
  "clear-outline-btn":   clearOutline,
  "match-height-btn":    matchHeight,
  "match-width-btn":     matchWidth,
  "match-size-btn":      matchHeightAndWidth,
  "distribute-resize-h-btn": distributeResizeHorizontal,
  "distribute-resize-v-btn": distributeResizeVertical,
  "stack-h-btn":         stackHorizontal,
  "stack-v-btn":         stackVertical,
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
  "bring-to-front-btn": bringShapesToFront,
  "send-to-back-btn": sendShapesToBack,
  "bring-forward-btn": bringShapesForward,
  "send-backward-btn": sendShapesBackward,
  "group-btn": groupSelectedShapes,
  "ungroup-btn": ungroupSelectedShapes,
  // Align & Group
  "align-center-group-btn":  alignCenterHAndGroup,
  "align-middle-group-btn":  alignMiddleVAndGroup,
  "center-middle-group-btn": centerMiddleAndGroup,
  "distribute-hv-group-btn": distributeHandVAndGroup,
  "distribute-h-group-btn":  distributeHAndGroup,
  "distribute-v-group-btn":  distributeVAndGroup,
  // Text Tools
  "default-textbox-format-btn": applyDefaultTextboxFormat,
  "merge-textboxes-btn": mergeTextBoxes,
  "split-textbox-btn": splitTextBoxByLines,
  "text-margins-none-btn": setTextMarginsNone,
  "text-margins-tight-btn": setTextMarginsTight,
  "text-margins-roomy-btn": setTextMarginsRoomy,
  "word-wrap-btn": toggleWordWrap,
  "autofit-off-btn": setTextAutofitOff,
  "autofit-text-btn": setTextAutofitTextToFitShape,
  "autofit-shape-btn": setTextAutofitShapeToFitText,
  "swap-text-btn": swapTextOnly,
  "insert-textbox-btn": insertTextBox,
  "insert-rectangle-btn": insertRectangle,
  "insert-arrow-btn": insertArrow,
  "insert-oval-btn": insertOval,
  "insert-line-btn": insertStraightLine,
  "insert-elbow-line-btn": insertElbowLine,
  "create-center-sticker-btn": createCenterSticker,
  "insert-rounded-rectangle-btn": insertRoundedRectangle,
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
  return Array.from(document.querySelectorAll<HTMLButtonElement>("button"));
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

function paletteRunner(target: PaletteTarget, color: string): ActionRunner {
  switch (target) {
    case "font":
      return () => applyFontColor(color);
    case "outline":
      return () => applyOutlineColor(color);
    case "fill":
      return () => applyFillColor(color);
  }
}

function renderPaletteColumn(
  containerId: string,
  target: PaletteTarget,
  colors: string[],
  rowSize: number,
  label: string,
  glyph: string,
  description: string
) {
  const container = document.getElementById(containerId);
  if (!container) {
    return;
  }

  const fragment = document.createDocumentFragment();

  colors.forEach((color, index) => {
    if (rowSize > 0 && index >= rowSize && index % rowSize === 0) {
      const spacer = document.createElement("div");
      spacer.className = "swatch-row-gap";
      spacer.setAttribute("aria-hidden", "true");
      fragment.appendChild(spacer);
    }

    const button = document.createElement("button");
    button.type = "button";
    button.className = "swatch-button";
    button.style.setProperty("--swatch", color);
    button.setAttribute("aria-label", `${label} color ${index + 1}`);
    button.title = `${label}\n${description}\n${color}`;
    button.dataset.paletteTarget = target;
    button.dataset.color = color;

    const swatch = document.createElement("span");
    swatch.className = "swatch-chip";
    swatch.setAttribute("aria-hidden", "true");
    button.appendChild(swatch);

    button.addEventListener("click", async () => {
      setBusy(true);
      try {
        const result = await paletteRunner(target, color)();
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

    fragment.appendChild(button);
  });

  container.appendChild(fragment);

  const glyphEl = document.querySelector<HTMLElement>(`[data-palette-glyph='${target}']`);
  if (glyphEl) {
    glyphEl.textContent = glyph;
    glyphEl.title = `${label}\n${description}`;
  }
}

async function setupPalette() {
  const palette = await getCurrentPresentationThemePalette();
  const sourceBadge = document.getElementById("palette-source");
  if (sourceBadge) {
    sourceBadge.textContent = palette.source === "deck" ? "Deck theme" : "Fallback palette";
  }

  renderPaletteColumn("font-palette", "font", palette.colors, palette.rowSize, "Font", "A", "Apply a color to the selected text.");
  renderPaletteColumn("outline-palette", "outline", palette.colors, palette.rowSize, "Outline", "O", "Apply a color to shape outlines.");
  renderPaletteColumn("fill-palette", "fill", palette.colors, palette.rowSize, "Fill", "F", "Apply a color to shape fills.");
}

Office.onReady(async () => {
  await setupPalette();
  applyNativeTooltips();

  Object.entries(ACTIONS).forEach(([id, runner]) => {
    wireAction(id, runner);
  });

  setStatus({
    type: "info",
    message: "Ready when you are.",
  });
});
