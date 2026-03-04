import type { ActionResult } from "../shared/shapeTools";
import { copyPosition, pastePosition, swapPositions } from "../shared/shapeTools";

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

export function copyPositionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, copyPosition);
}

export function pastePositionCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, pastePosition);
}

export function swapPositionsCommand(event: Office.AddinCommands.Event) {
  void withCommandEvent(event, swapPositions);
}

// Make them global for ExecuteFunction
(window as any).copyPosition = copyPositionCommand;
(window as any).pastePosition = pastePositionCommand;
(window as any).swapPositions = swapPositionsCommand;
