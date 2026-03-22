import type { ActionResult } from "../shared/shapeTools";
import {
  swapPositions,
  copyPositionAndSize,
  copyPositionOnly,
  copySizeOnly,
  pastePositionAndSize,
  pastePositionOnly,
  pasteSizeOnly,
  copyOutlineToClipboard,
  pasteOutlineFromClipboard,
  matchOutlineStyle,
  copyFillToClipboard,
  pasteFillFromClipboard,
  matchFillStyle,
  copyStyleToClipboard,
  pasteStyleFromClipboard,
  matchShapeStyle,
  clearFill,
  clearOutline,
  matchHeight,
  matchWidth,
  matchHeightAndWidth,
  alignLeft,
  alignCenterH,
  alignRight,
  alignTop,
  alignMiddleV,
  alignBottom,
  distributeH,
  distributeV,
  stretchToLeftEdge,
  stretchToRightEdge,
  stretchToTopEdge,
  stretchToBottomEdge,
  splitTextBoxByLines,
  removeTextMargins,
  disableTextAutofit,
  createCenterSticker,
  toggleDraftSticker,
  mergeTextBoxes,
  alignCenterHAndGroup,
  alignMiddleVAndGroup,
  centerMiddleAndGroup,
  distributeHAndGroup,
  distributeVAndGroup,
  openGridDialog,
  openWeekdayRangeDialog,
  openSelectedDeckDialog,
  openJolifyWebsite,
  moveToUnusedSection,
} from "../shared/shapeTools";

async function withCommandEvent(
  event: Office.AddinCommands.Event,
  runner: () => Promise<ActionResult>,
) {
  try {
    const result = await runner();
    if (result.type === "error") {
      console.error(result.message);
    } else if (result.type === "warning") {
      console.warn(result.message);
    } else {
      console.log(result.message);
    }
  } catch (error) {
    console.error(error);
  } finally {
    event.completed();
  }
}

export function swapPositionsCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, swapPositions);
}

// Position & Size
export function copyPositionAndSizeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyPositionAndSize);
}
export function copyPositionOnlyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyPositionOnly);
}
export function copySizeOnlyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copySizeOnly);
}
export function pastePositionAndSizeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pastePositionAndSize);
}
export function pastePositionOnlyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pastePositionOnly);
}
export function pasteSizeOnlyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pasteSizeOnly);
}

// Style
export function copyOutlineToClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyOutlineToClipboard);
}
export function pasteOutlineFromClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pasteOutlineFromClipboard);
}
export function matchOutlineStyleCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchOutlineStyle);
}
export function copyFillToClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyFillToClipboard);
}
export function pasteFillFromClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pasteFillFromClipboard);
}
export function matchFillStyleCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchFillStyle);
}
export function copyStyleToClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyStyleToClipboard);
}
export function pasteStyleFromClipboardCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pasteStyleFromClipboard);
}
export function matchShapeStyleCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchShapeStyle);
}
export function clearFillCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, clearFill);
}
export function clearOutlineCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, clearOutline);
}
export function matchHeightCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchHeight);
}
export function matchWidthCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchWidth);
}
export function matchHeightAndWidthCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, matchHeightAndWidth);
}

// Align
export function alignLeftCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignLeft);
}
export function alignCenterHCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignCenterH);
}
export function alignRightCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignRight);
}
export function alignTopCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignTop);
}
export function alignMiddleVCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignMiddleV);
}
export function alignBottomCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignBottom);
}
export function distributeHCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, distributeH);
}
export function distributeVCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, distributeV);
}
export function stretchToLeftEdgeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, stretchToLeftEdge);
}
export function stretchToRightEdgeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, stretchToRightEdge);
}
export function stretchToTopEdgeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, stretchToTopEdge);
}
export function stretchToBottomEdgeCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, stretchToBottomEdge);
}

// Text
export function splitTextBoxByLinesCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, splitTextBoxByLines);
}
export function removeTextMarginsCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, removeTextMargins);
}
export function disableTextAutofitCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, disableTextAutofit);
}
export function createCenterStickerCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, createCenterSticker);
}

// Branding
export function toggleDraftStickerCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, toggleDraftSticker);
}

// Phase 1 — new button-only commands
export function mergeTextBoxesCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, mergeTextBoxes);
}
export function alignCenterHAndGroupCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignCenterHAndGroup);
}
export function alignMiddleVAndGroupCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, alignMiddleVAndGroup);
}
export function centerMiddleAndGroupCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, centerMiddleAndGroup);
}
export function distributeHAndGroupCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, distributeHAndGroup);
}
export function distributeVAndGroupCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, distributeVAndGroup);
}
export function openGridDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openGridDialog);
}
export function openWeekdayRangeDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openWeekdayRangeDialog);
}
export function openSelectedDeckDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openSelectedDeckDialog);
}
export function openJolifyWebsiteCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openJolifyWebsite);
}
export function moveToUnusedSectionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, moveToUnusedSection);
}

// Signal to Office that the function file is ready
Office.onReady(() => {});

// Make them global for ExecuteFunction
(window as any).swapPositions = swapPositionsCommand;
(window as any).copyPositionAndSize = copyPositionAndSizeCommand;
(window as any).copyPositionOnly = copyPositionOnlyCommand;
(window as any).copySizeOnly = copySizeOnlyCommand;
(window as any).pastePositionAndSize = pastePositionAndSizeCommand;
(window as any).pastePositionOnly = pastePositionOnlyCommand;
(window as any).pasteSizeOnly = pasteSizeOnlyCommand;
(window as any).copyOutlineToClipboard = copyOutlineToClipboardCommand;
(window as any).pasteOutlineFromClipboard = pasteOutlineFromClipboardCommand;
(window as any).matchOutlineStyle = matchOutlineStyleCommand;
(window as any).copyFillToClipboard = copyFillToClipboardCommand;
(window as any).pasteFillFromClipboard = pasteFillFromClipboardCommand;
(window as any).matchFillStyle = matchFillStyleCommand;
(window as any).copyStyleToClipboard = copyStyleToClipboardCommand;
(window as any).pasteStyleFromClipboard = pasteStyleFromClipboardCommand;
(window as any).matchShapeStyle = matchShapeStyleCommand;
(window as any).clearFill = clearFillCommand;
(window as any).clearOutline = clearOutlineCommand;
(window as any).matchHeight = matchHeightCommand;
(window as any).matchWidth = matchWidthCommand;
(window as any).matchHeightAndWidth = matchHeightAndWidthCommand;
(window as any).alignLeft = alignLeftCommand;
(window as any).alignCenterH = alignCenterHCommand;
(window as any).alignRight = alignRightCommand;
(window as any).alignTop = alignTopCommand;
(window as any).alignMiddleV = alignMiddleVCommand;
(window as any).alignBottom = alignBottomCommand;
(window as any).distributeH = distributeHCommand;
(window as any).distributeV = distributeVCommand;
(window as any).stretchToLeftEdge = stretchToLeftEdgeCommand;
(window as any).stretchToRightEdge = stretchToRightEdgeCommand;
(window as any).stretchToTopEdge = stretchToTopEdgeCommand;
(window as any).stretchToBottomEdge = stretchToBottomEdgeCommand;
(window as any).splitTextBoxByLines = splitTextBoxByLinesCommand;
(window as any).removeTextMargins = removeTextMarginsCommand;
(window as any).disableTextAutofit = disableTextAutofitCommand;
(window as any).createCenterSticker = createCenterStickerCommand;
(window as any).toggleDraftSticker = toggleDraftStickerCommand;
(window as any).mergeTextBoxes = mergeTextBoxesCommand;
(window as any).alignCenterHAndGroup = alignCenterHAndGroupCommand;
(window as any).alignMiddleVAndGroup = alignMiddleVAndGroupCommand;
(window as any).centerMiddleAndGroup = centerMiddleAndGroupCommand;
(window as any).distributeHAndGroup = distributeHAndGroupCommand;
(window as any).distributeVAndGroup = distributeVAndGroupCommand;
(window as any).openGridDialog = openGridDialogCommand;
(window as any).openWeekdayRangeDialog = openWeekdayRangeDialogCommand;
(window as any).openSelectedDeckDialog = openSelectedDeckDialogCommand;
(window as any).openJolifyWebsite = openJolifyWebsiteCommand;
(window as any).moveToUnusedSection = moveToUnusedSectionCommand;
