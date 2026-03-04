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
