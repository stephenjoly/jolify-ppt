import type { ActionResult } from "../shared/shapeTools";
import {
  copyPosition,
  pastePosition,
  swapPositions,
  copyPositionAndSize,
  copyPositionOnly,
  copySizeOnly,
  pastePositionAndSize,
  pastePositionOnly,
  pasteSizeOnly,
  copyOutlineStyle,
  copyFillStyle,
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
  smartAnchorAlign,
  autoFontEqualizer,
  batchStyleApply,
  exportShapeMetadata,
  autoFlowText,
  normalizeConnectors,
  runAccessibilityCheck,
  convertTableToGantt,
  pasteAsGrid,
  openGridDialog,
  openColumnsDialog,
  openRowsDialog,
  openRenameDialog,
  openGanttDialog,
  openTimelineDialog,
  openSlideOutlineDialog,
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

// Position (legacy aliases — kept for backwards compatibility)
export function copyPositionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyPosition);
}
export function pastePositionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pastePosition);
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
export function smartAnchorAlignCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, smartAnchorAlign);
}
export function autoFontEqualizerCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, autoFontEqualizer);
}
export function batchStyleApplyCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, batchStyleApply);
}
export function exportShapeMetadataCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, exportShapeMetadata);
}
export function autoFlowTextCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, autoFlowText);
}
export function normalizeConnectorsCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, normalizeConnectors);
}
export function runAccessibilityCheckCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, runAccessibilityCheck);
}
export function convertTableToGanttCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, convertTableToGantt);
}
export function pasteAsGridCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pasteAsGrid);
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
export function openRenameDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openRenameDialog);
}
export function openGanttDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openGanttDialog);
}
export function openTimelineDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openTimelineDialog);
}
export function openSlideOutlineDialogCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, openSlideOutlineDialog);
}
export function moveToUnusedSectionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, moveToUnusedSection);
}

// Signal to Office that the function file is ready
Office.onReady(() => {});

// Make them global for ExecuteFunction
(window as any).copyPosition = copyPositionCommand;
(window as any).pastePosition = pastePositionCommand;
(window as any).swapPositions = swapPositionsCommand;
(window as any).copyPositionAndSize = copyPositionAndSizeCommand;
(window as any).copyPositionOnly = copyPositionOnlyCommand;
(window as any).copySizeOnly = copySizeOnlyCommand;
(window as any).pastePositionAndSize = pastePositionAndSizeCommand;
(window as any).pastePositionOnly = pastePositionOnlyCommand;
(window as any).pasteSizeOnly = pasteSizeOnlyCommand;
(window as any).copyOutlineStyle = copyOutlineStyleCommand;
(window as any).copyFillStyle = copyFillStyleCommand;
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
(window as any).smartAnchorAlign = smartAnchorAlignCommand;
(window as any).autoFontEqualizer = autoFontEqualizerCommand;
(window as any).batchStyleApply = batchStyleApplyCommand;
(window as any).exportShapeMetadata = exportShapeMetadataCommand;
(window as any).autoFlowText = autoFlowTextCommand;
(window as any).normalizeConnectors = normalizeConnectorsCommand;
(window as any).runAccessibilityCheck = runAccessibilityCheckCommand;
(window as any).convertTableToGantt = convertTableToGanttCommand;
(window as any).pasteAsGrid = pasteAsGridCommand;
(window as any).openGridDialog = openGridDialogCommand;
(window as any).openColumnsDialog = openColumnsDialogCommand;
(window as any).openRowsDialog = openRowsDialogCommand;
(window as any).openRenameDialog = openRenameDialogCommand;
(window as any).openGanttDialog = openGanttDialogCommand;
(window as any).openTimelineDialog = openTimelineDialogCommand;
(window as any).openSlideOutlineDialog = openSlideOutlineDialogCommand;
(window as any).moveToUnusedSection = moveToUnusedSectionCommand;
