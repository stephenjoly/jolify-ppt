import type { ActionResult } from "../shared/shapeTools";
import {
  swapPositions,
  copyPositionAndSize,
  copyPositionOnly,
  copySizeOnly,
  pastePositionAndSize,
  pastePositionOnly,
  pasteSizeOnly,
  copyOutlineStyle,
  copyFillStyle,
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
  splitTextBoxByLines,
  addDraftSticker,
  removeDraftSticker,
  mergeTextBoxes,
  alignCenterHAndGroup,
  alignMiddleVAndGroup,
  centerMiddleAndGroup,
  distributeHAndGroup,
  distributeVAndGroup,
  batchStyleApply,
  openGridDialog,
  openColumnsDialog,
  openRowsDialog,
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

// Match
export function copyOutlineStyleCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyOutlineStyle);
}
export function copyFillStyleCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyFillStyle);
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

// Text
export function splitTextBoxByLinesCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, splitTextBoxByLines);
}

// Branding
export function addDraftStickerCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, addDraftSticker);
}
export function removeDraftStickerCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, removeDraftSticker);
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
export function batchStyleApplyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, batchStyleApply);
}
export function openGridDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openGridDialog);
}
export function openColumnsDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openColumnsDialog);
}
export function openRowsDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openRowsDialog);
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
(window as any).copyOutlineStyle = copyOutlineStyleCommand;
(window as any).copyFillStyle = copyFillStyleCommand;
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
(window as any).splitTextBoxByLines = splitTextBoxByLinesCommand;
(window as any).addDraftSticker = addDraftStickerCommand;
(window as any).removeDraftSticker = removeDraftStickerCommand;
(window as any).mergeTextBoxes = mergeTextBoxesCommand;
(window as any).alignCenterHAndGroup = alignCenterHAndGroupCommand;
(window as any).alignMiddleVAndGroup = alignMiddleVAndGroupCommand;
(window as any).centerMiddleAndGroup = centerMiddleAndGroupCommand;
(window as any).distributeHAndGroup = distributeHAndGroupCommand;
(window as any).distributeVAndGroup = distributeVAndGroupCommand;
(window as any).batchStyleApply = batchStyleApplyCommand;
(window as any).openGridDialog = openGridDialogCommand;
(window as any).openColumnsDialog = openColumnsDialogCommand;
(window as any).openRowsDialog = openRowsDialogCommand;
(window as any).moveToUnusedSection = moveToUnusedSectionCommand;
