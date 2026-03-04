export type ShapePosition = {
  left: number;
  top: number;
};

export type ShapeSize = {
  width: number;
  height: number;
};

export type ActionResult = {
  message: string;
  type: "success" | "info" | "warning" | "error";
};

let savedPosition: ShapePosition | null = null;
let savedSize: ShapeSize | null = null;

async function getSelectedShapes(context: PowerPoint.RequestContext) {
  const selection = context.presentation.getSelectedShapes();
  selection.load("items");
  await context.sync();
  return selection.items;
}

export async function copyPositionOnly(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its position.",
      };
    }

    const shape = shapes[0];
    shape.load("left,top");
    await context.sync();

    savedPosition = {
      left: shape.left,
      top: shape.top,
    };

    return {
      type: "success",
      message: "Saved position from the first selected shape.",
    };
  });
}

export async function copySizeOnly(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its size.",
      };
    }

    const shape = shapes[0];
    shape.load("width,height");
    await context.sync();

    savedSize = {
      width: shape.width,
      height: shape.height,
    };

    return {
      type: "success",
      message: "Saved size from the first selected shape.",
    };
  });
}

export async function copyPositionAndSize(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its position & size.",
      };
    }

    const shape = shapes[0];
    shape.load("left,top,width,height");
    await context.sync();

    savedPosition = {
      left: shape.left,
      top: shape.top,
    };

    savedSize = {
      width: shape.width,
      height: shape.height,
    };

    return {
      type: "success",
      message: "Saved position & size from the first selected shape.",
    };
  });
}

export function copyPosition(): Promise<ActionResult> {
  return copyPositionAndSize();
}

export async function pastePositionOnly(): Promise<ActionResult> {
  if (!savedPosition) {
    return {
      type: "warning",
      message: "Copy a shape's position before pasting.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select one or more shapes to paste the saved position.",
      };
    }

    shapes.forEach((shape) => {
      shape.left = savedPosition!.left;
      shape.top = savedPosition!.top;
    });

    await context.sync();

    return {
      type: "success",
      message: `Pasted saved position to ${shapes.length} shape(s).`,
    };
  });
}

export async function pasteSizeOnly(): Promise<ActionResult> {
  if (!savedSize) {
    return {
      type: "warning",
      message: "Copy a shape's size before pasting.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select one or more shapes to paste the saved size.",
      };
    }

    shapes.forEach((shape) => {
      shape.width = savedSize!.width;
      shape.height = savedSize!.height;
    });

    await context.sync();

    return {
      type: "success",
      message: `Pasted saved size to ${shapes.length} shape(s).`,
    };
  });
}

export async function pastePositionAndSize(): Promise<ActionResult> {
  if (!savedPosition || !savedSize) {
    return {
      type: "warning",
      message: "Copy position & size before pasting both.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select one or more shapes to paste the saved position & size.",
      };
    }

    shapes.forEach((shape) => {
      shape.left = savedPosition!.left;
      shape.top = savedPosition!.top;
      shape.width = savedSize!.width;
      shape.height = savedSize!.height;
    });

    await context.sync();

    return {
      type: "success",
      message: `Pasted saved position & size to ${shapes.length} shape(s).`,
    };
  });
}

export function pastePosition(): Promise<ActionResult> {
  return pastePositionAndSize();
}

export async function swapPositions(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length !== 2) {
      return {
        type: "warning",
        message: "Select exactly two shapes to swap their positions.",
      };
    }

    const [first, second] = shapes;
    first.load("left,top,width,height");
    second.load("left,top,width,height");
    await context.sync();

    const firstPos: ShapePosition & ShapeSize = {
      left: first.left,
      top: first.top,
      width: first.width,
      height: first.height,
    };
    const secondPos: ShapePosition & ShapeSize = {
      left: second.left,
      top: second.top,
      width: second.width,
      height: second.height,
    };

    first.left = secondPos.left;
    first.top = secondPos.top;
    // first.width = secondPos.width;
    // first.height = secondPos.height;

    second.left = firstPos.left;
    second.top = firstPos.top;
    // second.width = firstPos.width;
    // second.height = firstPos.height;

    await context.sync();

    return {
      type: "success",
      message: "Swapped positions of the two selected shapes.",
    };
  });
}

export function getSavedPosition(): ShapePosition | null {
  return savedPosition;
}

export function clearSavedPosition(): void {
  savedPosition = null;
}

export function getSavedSize(): ShapeSize | null {
  return savedSize;
}

export function clearSavedSize(): void {
  savedSize = null;
}

const DRAFT_STICKER_NAME = "__agetnoon_draft_sticker__";
const DRAFT_TEXT = "DRAFT";
const DEFAULT_STICKER_WIDTH = 140;
const DEFAULT_STICKER_HEIGHT = 40;
const DEFAULT_RIGHT_OFFSET = 40;
const DEFAULT_TOP_OFFSET = 16;

async function getPrimarySlideMaster(context: PowerPoint.RequestContext) {
  const masters = context.presentation.slideMasters;
  masters.load("items");
  await context.sync();
  return masters.items[0] ?? null;
}

async function loadShapeNames(shapes: PowerPoint.ShapeCollection) {
  shapes.load("items");
  await shapes.context.sync();
  shapes.items.forEach((shape) => shape.load("name"));
  await shapes.context.sync();
}

export async function addDraftSticker(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const master = await getPrimarySlideMaster(context);
    if (!master) {
      return {
        type: "warning",
        message: "No slide master found; unable to add draft sticker.",
      };
    }

    const shapes = master.shapes;
    await loadShapeNames(shapes);

    const existing = shapes.items.find((shape) => shape.name === DRAFT_STICKER_NAME);
    if (existing) {
      return {
        type: "info",
        message: "Draft sticker already present on the slide master.",
      };
    }

    const sticker = shapes.addTextBox(DRAFT_TEXT, {
      width: DEFAULT_STICKER_WIDTH,
      height: DEFAULT_STICKER_HEIGHT,
      top: DEFAULT_TOP_OFFSET,
      left: DEFAULT_RIGHT_OFFSET,
    });

    sticker.name = DRAFT_STICKER_NAME;
    sticker.textFrame.textRange.font.color = "#ffffff";
    sticker.textFrame.textRange.font.bold = true;
    sticker.textFrame.textRange.paragraphFormat.alignment = "Center";
    sticker.width = DEFAULT_STICKER_WIDTH;
    sticker.height = DEFAULT_STICKER_HEIGHT;
    sticker.top = DEFAULT_TOP_OFFSET;
    // Approximate the right edge by moving the sticker after width is set.
    sticker.left = Math.max(DEFAULT_RIGHT_OFFSET, sticker.left);
    sticker.fill.setSolidColor("#a4262c");
    sticker.lineFormat.visible = false;

    await context.sync();

    // Move sticker closer to the top-right corner using slide width heuristic (default 960 points).
    // Without a direct API for slide dimensions, assume widescreen width and clamp to non-negative.
    const assumedSlideWidth = 960;
    sticker.left = Math.max(
      0,
      assumedSlideWidth - DEFAULT_RIGHT_OFFSET - sticker.width,
    );

    await context.sync();

    return {
      type: "success",
      message: "Draft sticker added to the slide master.",
    };
  });
}

export async function removeDraftSticker(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const master = await getPrimarySlideMaster(context);
    if (!master) {
      return {
        type: "warning",
        message: "No slide master found; nothing to remove.",
      };
    }

    const shapes = master.shapes;
    await loadShapeNames(shapes);

    const matches = shapes.items.filter((shape) => shape.name === DRAFT_STICKER_NAME);
    if (matches.length === 0) {
      return {
        type: "info",
        message: "Draft sticker is already removed.",
      };
    }

    matches.forEach((shape) => shape.delete());
    await context.sync();

    return {
      type: "success",
      message: "Removed draft sticker from the slide master.",
    };
  });
}

const SLIDE_WIDTH = 960;  // points, widescreen default (same assumption as addDraftSticker)
const SLIDE_HEIGHT = 540; // points, widescreen default

type AlignType = "left" | "centerH" | "right" | "top" | "middleV" | "bottom" | "distributeH" | "distributeV";

async function alignShapes(type: AlignType): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 1) {
      return { type: "warning", message: "Select at least one shape to align." };
    }

    const isDistribute = type === "distributeH" || type === "distributeV";
    if (isDistribute && shapes.length < 3) {
      return { type: "warning", message: "Select at least 3 shapes to distribute." };
    }

    shapes.forEach((s) => s.load("left,top,width,height"));
    await context.sync();

    if (shapes.length === 1) {
      const s = shapes[0];
      switch (type) {
        case "left":    s.left = 0; break;
        case "centerH": s.left = (SLIDE_WIDTH - s.width) / 2; break;
        case "right":   s.left = SLIDE_WIDTH - s.width; break;
        case "top":     s.top = 0; break;
        case "middleV": s.top = (SLIDE_HEIGHT - s.height) / 2; break;
        case "bottom":  s.top = SLIDE_HEIGHT - s.height; break;
      }
    } else {
      switch (type) {
        case "left": {
          const anchor = Math.min(...shapes.map((s) => s.left));
          shapes.forEach((s) => { s.left = anchor; });
          break;
        }
        case "centerH": {
          const minLeft = Math.min(...shapes.map((s) => s.left));
          const maxRight = Math.max(...shapes.map((s) => s.left + s.width));
          const center = (minLeft + maxRight) / 2;
          shapes.forEach((s) => { s.left = center - s.width / 2; });
          break;
        }
        case "right": {
          const anchor = Math.max(...shapes.map((s) => s.left + s.width));
          shapes.forEach((s) => { s.left = anchor - s.width; });
          break;
        }
        case "top": {
          const anchor = Math.min(...shapes.map((s) => s.top));
          shapes.forEach((s) => { s.top = anchor; });
          break;
        }
        case "middleV": {
          const minTop = Math.min(...shapes.map((s) => s.top));
          const maxBottom = Math.max(...shapes.map((s) => s.top + s.height));
          const middle = (minTop + maxBottom) / 2;
          shapes.forEach((s) => { s.top = middle - s.height / 2; });
          break;
        }
        case "bottom": {
          const anchor = Math.max(...shapes.map((s) => s.top + s.height));
          shapes.forEach((s) => { s.top = anchor - s.height; });
          break;
        }
        case "distributeH": {
          const sorted = [...shapes].sort((a, b) => a.left - b.left);
          const span = sorted[sorted.length - 1].left + sorted[sorted.length - 1].width - sorted[0].left;
          const totalWidth = sorted.reduce((sum, s) => sum + s.width, 0);
          const gap = (span - totalWidth) / (sorted.length - 1);
          let cursor = sorted[0].left;
          sorted.forEach((s) => { s.left = cursor; cursor += s.width + gap; });
          break;
        }
        case "distributeV": {
          const sorted = [...shapes].sort((a, b) => a.top - b.top);
          const span = sorted[sorted.length - 1].top + sorted[sorted.length - 1].height - sorted[0].top;
          const totalHeight = sorted.reduce((sum, s) => sum + s.height, 0);
          const gap = (span - totalHeight) / (sorted.length - 1);
          let cursor = sorted[0].top;
          sorted.forEach((s) => { s.top = cursor; cursor += s.height + gap; });
          break;
        }
      }
    }

    await context.sync();

    const verb = isDistribute ? "Distributed" : "Aligned";
    return {
      type: "success",
      message: `${verb} ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function copyFillStyle(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — fill is copied from the first to the rest.",
      };
    }

    const source = shapes[0];
    source.fill.load("type,foregroundColor,transparency");
    await context.sync();

    const fillType = source.fill.type;

    if (fillType !== "NoFill" && fillType !== "Solid") {
      return {
        type: "warning",
        message: "Only solid color and no-fill can be copied. Gradient, pattern, and picture fills aren't supported by the Office.js API.",
      };
    }

    const { foregroundColor, transparency } = source.fill;

    shapes.slice(1).forEach((shape) => {
      if (fillType === "NoFill") {
        shape.fill.clear();
      } else {
        shape.fill.setSolidColor(foregroundColor);
        shape.fill.transparency = transparency;
      }
    });

    await context.sync();

    return {
      type: "success",
      message: `Copied fill to ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function matchHeight(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — height is copied from the first to the rest.",
      };
    }

    shapes[0].load("height");
    await context.sync();

    const { height } = shapes[0];
    shapes.slice(1).forEach((shape) => { shape.height = height; });
    await context.sync();

    return {
      type: "success",
      message: `Matched height of ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function matchWidth(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — width is copied from the first to the rest.",
      };
    }

    shapes[0].load("width");
    await context.sync();

    const { width } = shapes[0];
    shapes.slice(1).forEach((shape) => { shape.width = width; });
    await context.sync();

    return {
      type: "success",
      message: `Matched width of ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function matchHeightAndWidth(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — size is copied from the first to the rest.",
      };
    }

    shapes[0].load("height,width");
    await context.sync();

    const { height, width } = shapes[0];
    shapes.slice(1).forEach((shape) => {
      shape.height = height;
      shape.width = width;
    });
    await context.sync();

    return {
      type: "success",
      message: `Matched size of ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function copyOutlineStyle(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — outline is copied from the first to the rest.",
      };
    }

    const source = shapes[0];
    source.lineFormat.load("color,dashStyle,style,transparency,visible,weight");
    await context.sync();

    const { color, dashStyle, style, transparency, visible, weight } = source.lineFormat;

    shapes.slice(1).forEach((shape) => {
      shape.lineFormat.visible = visible;
      shape.lineFormat.color = color;
      shape.lineFormat.dashStyle = dashStyle;
      shape.lineFormat.style = style;
      shape.lineFormat.transparency = transparency;
      shape.lineFormat.weight = weight;
    });

    await context.sync();

    return {
      type: "success",
      message: `Copied outline to ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export const alignLeft    = (): Promise<ActionResult> => alignShapes("left");
export const alignCenterH = (): Promise<ActionResult> => alignShapes("centerH");
export const alignRight   = (): Promise<ActionResult> => alignShapes("right");
export const alignTop     = (): Promise<ActionResult> => alignShapes("top");
export const alignMiddleV = (): Promise<ActionResult> => alignShapes("middleV");
export const alignBottom  = (): Promise<ActionResult> => alignShapes("bottom");
export const distributeH  = (): Promise<ActionResult> => alignShapes("distributeH");
export const distributeV  = (): Promise<ActionResult> => alignShapes("distributeV");

export async function splitTextBoxByLines(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length !== 1) {
      return {
        type: "warning",
        message: "Select exactly one text box to split.",
      };
    }

    const shape = shapes[0];
    shape.load("left,top,width,height");
    shape.textFrame.textRange.load("text");
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return {
        type: "error",
        message: "Could not determine the current slide.",
      };
    }

    const fullText = shape.textFrame.textRange.text;
    const lines = fullText
      ? fullText.split(/\r\n|\r|\n/).filter((line) => line.trim() !== "")
      : [];

    if (lines.length <= 1) {
      return {
        type: "info",
        message:
          lines.length === 0
            ? "The text box is empty."
            : "The text box has only one line; nothing to split.",
      };
    }

    const { left, top, width, height } = shape;
    const lineHeight = height / lines.length;
    const slide = selectedSlides.items[0];

    lines.forEach((line, index) => {
      slide.shapes.addTextBox(line, {
        left,
        top: top + index * lineHeight,
        width,
        height: lineHeight,
      });
    });

    shape.delete();
    await context.sync();

    return {
      type: "success",
      message: `Split into ${lines.length} text boxes.`,
    };
  });
}

export async function removeAllComments(): Promise<ActionResult> {
  return {
    type: "warning",
    message:
      "PowerPoint's Office.js APIs do not currently expose slide comments, so they can't be read or removed programmatically.",
  };
}

export async function removeAllSpeakerNotes(): Promise<ActionResult> {
  return {
    type: "warning",
    message:
      "PowerPoint's Office.js APIs do not provide access to speaker notes, so they can't be exported or cleared programmatically.",
  };
}
