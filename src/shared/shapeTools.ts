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

const DRAFT_STICKER_NAME = "__jolify_draft_sticker__";
const STICKER_SIZE = 81;  // points — 2.87 cm
const STICKER_IMAGE_URL = "./assets/draft-sticker.png";

async function fetchImageAsBase64(url: string): Promise<string> {
  const response = await fetch(url);
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve((reader.result as string).split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export async function addDraftSticker(): Promise<ActionResult> {
  let imageBase64: string;
  try {
    imageBase64 = await fetchImageAsBase64(STICKER_IMAGE_URL);
  } catch {
    return { type: "error", message: "Could not load the draft sticker image." };
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) {
      return { type: "warning", message: "No slides found." };
    }

    let added = 0;
    let skipped = 0;

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();
      shapes.items.forEach((s) => s.load("name"));
      await context.sync();

      if (shapes.items.some((s) => s.name === DRAFT_STICKER_NAME)) {
        skipped++;
        continue;
      }

      const sticker = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 879,  // 31.04 cm in pt
        top: -1,    // -0.04 cm in pt
        width: STICKER_SIZE,
        height: STICKER_SIZE,
      });
      sticker.name = DRAFT_STICKER_NAME;
      sticker.fill.setImage(imageBase64);
      sticker.lineFormat.visible = false;

      await context.sync();
      added++;
    }

    if (added === 0) {
      return { type: "info", message: "All slides already have a DRAFT sticker." };
    }

    const skippedNote = skipped > 0 ? ` (${skipped} already had it)` : "";
    return {
      type: "success",
      message: `Added DRAFT sticker to ${added} slide${added !== 1 ? "s" : ""}.${skippedNote}`,
    };
  });
}

export async function removeDraftSticker(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    let removed = 0;

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();
      shapes.items.forEach((s) => s.load("name"));
      await context.sync();

      shapes.items
        .filter((s) => s.name === DRAFT_STICKER_NAME)
        .forEach((s) => { s.delete(); removed++; });
    }

    await context.sync();

    if (removed === 0) {
      return { type: "info", message: "No DRAFT stickers found." };
    }

    return {
      type: "success",
      message: `Removed DRAFT sticker from ${removed} slide${removed !== 1 ? "s" : ""}.`,
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

// ─────────────────────────────────────────────────────────────────
// Phase 1 — Button-only features
// ─────────────────────────────────────────────────────────────────

export async function mergeTextBoxes(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 text boxes to merge." };
    }

    shapes.forEach((s) => {
      s.load("left,top,width,height");
      s.textFrame.textRange.load("text");
    });
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const texts = shapes.map((s) => s.textFrame.textRange.text?.trim() ?? "").filter(Boolean);
    const merged = texts.join("\n");

    const minLeft = Math.min(...shapes.map((s) => s.left));
    const minTop = Math.min(...shapes.map((s) => s.top));
    const maxRight = Math.max(...shapes.map((s) => s.left + s.width));
    const maxBottom = Math.max(...shapes.map((s) => s.top + s.height));

    const slide = selectedSlides.items[0];
    slide.shapes.addTextBox(merged, {
      left: minLeft,
      top: minTop,
      width: maxRight - minLeft,
      height: maxBottom - minTop,
    });

    shapes.forEach((s) => s.delete());
    await context.sync();

    return { type: "success", message: `Merged ${shapes.length} text boxes into one.` };
  });
}

async function alignAndGroup(type: AlignType): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 shapes to align and group." };
    }

    const isDistribute = type === "distributeH" || type === "distributeV";
    if (isDistribute && shapes.length < 3) {
      return { type: "warning", message: "Select at least 3 shapes to distribute." };
    }

    shapes.forEach((s) => s.load("left,top,width,height"));
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    // Apply alignment
    switch (type) {
      case "centerH": {
        const minLeft = Math.min(...shapes.map((s) => s.left));
        const maxRight = Math.max(...shapes.map((s) => s.left + s.width));
        const center = (minLeft + maxRight) / 2;
        shapes.forEach((s) => { s.left = center - s.width / 2; });
        break;
      }
      case "middleV": {
        const minTop = Math.min(...shapes.map((s) => s.top));
        const maxBottom = Math.max(...shapes.map((s) => s.top + s.height));
        const middle = (minTop + maxBottom) / 2;
        shapes.forEach((s) => { s.top = middle - s.height / 2; });
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

    await context.sync();

    const slide = selectedSlides.items[0];
    slide.shapes.addGroup(shapes);
    await context.sync();

    const verb = isDistribute ? "Distributed and grouped" : "Aligned and grouped";
    return { type: "success", message: `${verb} ${shapes.length} shapes.` };
  });
}

export const alignCenterHAndGroup  = (): Promise<ActionResult> => alignAndGroup("centerH");
export const alignMiddleVAndGroup  = (): Promise<ActionResult> => alignAndGroup("middleV");
export const distributeHAndGroup   = (): Promise<ActionResult> => alignAndGroup("distributeH");
export const distributeVAndGroup   = (): Promise<ActionResult> => alignAndGroup("distributeV");

export async function centerMiddleAndGroup(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 shapes to align and group." };
    }

    shapes.forEach((s) => s.load("left,top,width,height"));
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const minLeft = Math.min(...shapes.map((s) => s.left));
    const maxRight = Math.max(...shapes.map((s) => s.left + s.width));
    const center = (minLeft + maxRight) / 2;
    shapes.forEach((s) => { s.left = center - s.width / 2; });

    const minTop = Math.min(...shapes.map((s) => s.top));
    const maxBottom = Math.max(...shapes.map((s) => s.top + s.height));
    const middle = (minTop + maxBottom) / 2;
    shapes.forEach((s) => { s.top = middle - s.height / 2; });

    await context.sync();

    selectedSlides.items[0].shapes.addGroup(shapes);
    await context.sync();

    return { type: "success", message: `Center+middle aligned and grouped ${shapes.length} shapes.` };
  });
}

export async function smartAnchorAlign(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 shapes — the first is the anchor." };
    }

    shapes.forEach((s) => s.load("left,top,width,height"));
    await context.sync();

    const anchor = shapes[0];
    const anchorCenterX = anchor.left + anchor.width / 2;
    const anchorCenterY = anchor.top + anchor.height / 2;

    shapes.slice(1).forEach((s) => {
      s.left = anchorCenterX - s.width / 2;
      s.top = anchorCenterY - s.height / 2;
    });

    await context.sync();

    return {
      type: "success",
      message: `Aligned ${shapes.length - 1} shape(s) to the anchor's center.`,
    };
  });
}

export async function autoFontEqualizer(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 text boxes to equalize font sizes." };
    }

    shapes.forEach((s) => s.textFrame.textRange.font.load("size"));
    await context.sync();

    const sizes = shapes
      .map((s) => s.textFrame.textRange.font.size)
      .filter((sz) => sz != null && sz > 0);

    if (sizes.length === 0) {
      return { type: "warning", message: "No font sizes found on selected shapes." };
    }

    const minSize = Math.min(...sizes);
    shapes.forEach((s) => { s.textFrame.textRange.font.size = minSize; });
    await context.sync();

    return {
      type: "success",
      message: `Set font size to ${minSize}pt on ${shapes.length} shapes.`,
    };
  });
}

export async function batchStyleApply(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — style is copied from the first to the rest.",
      };
    }

    const source = shapes[0];
    source.fill.load("type,foregroundColor,transparency");
    source.lineFormat.load("color,dashStyle,style,transparency,visible,weight");
    source.textFrame.textRange.font.load("bold,color,italic,size,underline");
    await context.sync();

    const fillType = source.fill.type;
    const { foregroundColor: fillColor, transparency: fillTransp } = source.fill;
    const { color: lineColor, dashStyle, style: lineStyle, transparency: lineTransp, visible, weight } = source.lineFormat;
    const { bold, color: fontColor, italic, size: fontSize, underline } = source.textFrame.textRange.font;

    shapes.slice(1).forEach((shape) => {
      // Fill
      if (fillType === "NoFill") {
        shape.fill.clear();
      } else if (fillType === "Solid") {
        shape.fill.setSolidColor(fillColor);
        shape.fill.transparency = fillTransp;
      }

      // Outline
      shape.lineFormat.visible = visible;
      if (visible) {
        shape.lineFormat.color = lineColor;
        shape.lineFormat.dashStyle = dashStyle;
        shape.lineFormat.style = lineStyle;
        shape.lineFormat.transparency = lineTransp;
        shape.lineFormat.weight = weight;
      }

      // Font
      shape.textFrame.textRange.font.bold = bold;
      shape.textFrame.textRange.font.color = fontColor;
      shape.textFrame.textRange.font.italic = italic;
      if (fontSize) shape.textFrame.textRange.font.size = fontSize;
      shape.textFrame.textRange.font.underline = underline;
    });

    await context.sync();

    return {
      type: "success",
      message: `Applied style from first shape to ${shapes.length - 1} shape(s).`,
    };
  });
}

export async function exportShapeMetadata(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const data: object[] = [];

    for (let slideIdx = 0; slideIdx < slides.items.length; slideIdx++) {
      const slide = slides.items[slideIdx];
      const shapeColl = slide.shapes;
      shapeColl.load("items");
      await context.sync();
      shapeColl.items.forEach((s) => s.load("name,left,top,width,height,type"));
      await context.sync();

      shapeColl.items.forEach((s) => {
        data.push({
          slide: slideIdx + 1,
          name: s.name,
          type: s.type,
          left: Math.round(s.left),
          top: Math.round(s.top),
          width: Math.round(s.width),
          height: Math.round(s.height),
        });
      });
    }

    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "shape-metadata.json";
    anchor.click();
    URL.revokeObjectURL(url);

    return {
      type: "success",
      message: `Exported metadata for ${data.length} shapes across ${slides.items.length} slide(s).`,
    };
  });
}

export async function autoFlowText(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return { type: "warning", message: "Select at least 2 text boxes to flow text across." };
    }

    shapes.forEach((s) => s.textFrame.textRange.load("text"));
    await context.sync();

    // Collect all words from all selected shapes
    const allText = shapes.map((s) => s.textFrame.textRange.text ?? "").join(" ");
    const words = allText.split(/\s+/).filter(Boolean);

    if (words.length === 0) {
      return { type: "info", message: "No text found in selected shapes." };
    }

    const perShape = Math.ceil(words.length / shapes.length);
    shapes.forEach((s, i) => {
      const chunk = words.slice(i * perShape, (i + 1) * perShape).join(" ");
      s.textFrame.textRange.text = chunk;
    });

    await context.sync();

    return {
      type: "success",
      message: `Redistributed ${words.length} word(s) across ${shapes.length} shapes.`,
    };
  });
}

export async function normalizeConnectors(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return { type: "warning", message: "Select at least one connector shape." };
    }

    shapes.forEach((s) => s.load("type"));
    await context.sync();

    const connectors = shapes.filter((s) => s.type === "Line");

    if (connectors.length === 0) {
      return { type: "info", message: "No connector/line shapes found in the selection." };
    }

    // Set connectors to use solid line style (normalize appearance)
    connectors.forEach((s) => {
      s.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      s.lineFormat.style = PowerPoint.ShapeLineStyle.single;
    });

    await context.sync();

    return {
      type: "success",
      message: `Normalized ${connectors.length} connector(s) to straight solid style.`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 1 — Diagnostics
// ─────────────────────────────────────────────────────────────────

export async function runAccessibilityCheck(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const issues: string[] = [];

    for (let slideIdx = 0; slideIdx < slides.items.length; slideIdx++) {
      const slide = slides.items[slideIdx];
      const shapeColl = slide.shapes;
      shapeColl.load("items");
      await context.sync();
      shapeColl.items.forEach((s) => s.load("name,type,altTextTitle,altTextDescription"));
      await context.sync();

      for (const shape of shapeColl.items) {
        const hasAltText = (shape.altTextTitle?.trim() || shape.altTextDescription?.trim());
        const isImage = shape.type === "Picture" || shape.type === "Media";

        if (isImage && !hasAltText) {
          issues.push(`Slide ${slideIdx + 1}: "${shape.name}" — image missing alt text`);
        }
      }
    }

    if (issues.length === 0) {
      return { type: "success", message: "No accessibility issues found." };
    }

    return {
      type: "warning",
      message: `Found ${issues.length} issue(s):\n${issues.slice(0, 5).join("\n")}${issues.length > 5 ? `\n…and ${issues.length - 5} more` : ""}`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 2 — Dialog-based helpers
// ─────────────────────────────────────────────────────────────────

export function openDialog<T>(relativeUrl: string): Promise<T | null> {
  return new Promise((resolve) => {
    const url = `${window.location.origin}${window.location.pathname.replace(/\/[^/]+$/, "/")}${relativeUrl}`;
    let dialog: Office.Dialog;

    Office.context.ui.displayDialogAsync(
      url,
      { height: 50, width: 40, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          resolve(null);
          return;
        }

        dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args: any) => {
          dialog.close();
          try {
            const parsed = JSON.parse(args.message) as T;
            resolve(parsed);
          } catch {
            resolve(null);
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          resolve(null);
        });
      },
    );
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 2 — Layout builders (grid / columns / rows)
// ─────────────────────────────────────────────────────────────────

type GapPreset = "none" | "small" | "medium" | "large";
type SizePreset = "small" | "medium" | "large" | "full";

const GAP_MAP: Record<GapPreset, number> = { none: 0, small: 10, medium: 20, large: 36 };
const SIZE_FRACTION: Record<SizePreset, number> = { small: 0.20, medium: 0.33, large: 0.50, full: 1.0 };

export type GridParams = {
  rows: number;
  cols: number;
  gapPreset: GapPreset;
  sizePreset: SizePreset;
};

export async function createGrid(params: GridParams): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const { rows, cols, gapPreset, sizePreset } = params;

    if (rows < 1 || cols < 1) {
      return { type: "warning", message: "Rows and columns must each be at least 1." };
    }

    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const gap = GAP_MAP[gapPreset];
    const fraction = SIZE_FRACTION[sizePreset];

    const availableW = SLIDE_WIDTH * fraction;
    const availableH = SLIDE_HEIGHT * fraction;

    const cellWidth = (availableW - (cols - 1) * gap) / cols;
    const cellHeight = (availableH - (rows - 1) * gap) / rows;

    const offsetX = (SLIDE_WIDTH - availableW) / 2;
    const offsetY = (SLIDE_HEIGHT - availableH) / 2;

    const slide = selectedSlides.items[0];

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: offsetX + c * (cellWidth + gap),
          top: offsetY + r * (cellHeight + gap),
          width: cellWidth,
          height: cellHeight,
        });
      }
    }

    await context.sync();

    return {
      type: "success",
      message: `Created ${rows}×${cols} grid (${rows * cols} shapes).`,
    };
  });
}

export type ColumnsParams = {
  count: number;
  gapPreset: GapPreset;
  sizePreset: SizePreset;
};

export async function createColumns(params: ColumnsParams): Promise<ActionResult> {
  return createGrid({ rows: 1, cols: params.count, gapPreset: params.gapPreset, sizePreset: params.sizePreset });
}

export type RowsParams = {
  count: number;
  gapPreset: GapPreset;
  sizePreset: SizePreset;
};

export async function createRows(params: RowsParams): Promise<ActionResult> {
  return createGrid({ rows: params.count, cols: 1, gapPreset: params.gapPreset, sizePreset: params.sizePreset });
}

export type RenameParams = {
  template: string;
};

export async function batchRenameShapes(params: RenameParams): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return { type: "warning", message: "Select at least one shape to rename." };
    }

    shapes.forEach((s, i) => {
      s.name = params.template
        .replace(/\{n\}/g, String(i + 1))
        .replace(/\{N\}/g, String(i + 1))
        .replace(/\{i\}/g, String(i))
        .replace(/\{I\}/g, String(i));
    });

    await context.sync();

    return {
      type: "success",
      message: `Renamed ${shapes.length} shape(s) using template "${params.template}".`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 2 — Dialog wrappers (open dialog then execute)
// ─────────────────────────────────────────────────────────────────

export async function openGridDialog(): Promise<ActionResult> {
  const params = await openDialog<GridParams>("dialogs/grid-builder.html");
  if (!params) return { type: "info", message: "Grid creation cancelled." };
  return createGrid(params);
}

export async function openColumnsDialog(): Promise<ActionResult> {
  const params = await openDialog<ColumnsParams>("dialogs/grid-builder.html?mode=columns");
  if (!params) return { type: "info", message: "Column creation cancelled." };
  return createColumns(params);
}

export async function openRowsDialog(): Promise<ActionResult> {
  const params = await openDialog<RowsParams>("dialogs/grid-builder.html?mode=rows");
  if (!params) return { type: "info", message: "Row creation cancelled." };
  return createRows(params);
}

export async function openRenameDialog(): Promise<ActionResult> {
  const params = await openDialog<RenameParams>("dialogs/rename-shapes.html");
  if (!params) return { type: "info", message: "Rename cancelled." };
  return batchRenameShapes(params);
}

// ─────────────────────────────────────────────────────────────────
// Phase 3 — Complex builders
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
// Gantt Chart — types & builder
// ─────────────────────────────────────────────────────────────────

const GANTT_TAG = "__jolify_gantt__";

export type GanttTaskBar = {
  label: string;
  start: string;  // ISO date "YYYY-MM-DD"
  end: string;    // ISO date "YYYY-MM-DD"
  color?: string; // hex, e.g. "#6B9E6B"
};

export type GanttRowEntry = {
  activity: string;
  team?: string;
  tasks: GanttTaskBar[];
  note?: string;
};

export type GanttChartParams = {
  title?: string;
  projectStart: string;
  projectEnd: string;
  granularity?: "week" | "biweek";
  todayDate?: string;
  rows: GanttRowEntry[];
};

// Keep old alias so convertTableToGantt compiles until updated below
type GanttTask = GanttTaskBar & { name: string };
type GanttParams = { tasks: GanttTask[]; projectStart: string; projectEnd: string };

export async function buildGanttChart(params: GanttChartParams): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    if (!params.rows || params.rows.length === 0) {
      return { type: "warning", message: "No rows provided for the Gantt chart." };
    }

    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const slide = selectedSlides.items[0];

    // Remove any existing Gantt shapes on this slide
    const allShapes = slide.shapes;
    allShapes.load("items");
    await context.sync();
    allShapes.items.forEach((s) => s.load("name"));
    await context.sync();
    allShapes.items
      .filter((s) => s.name?.startsWith(GANTT_TAG))
      .forEach((s) => s.delete());
    await context.sync();

    // --- Week columns ---
    const granDays = params.granularity === "biweek" ? 14 : 7;
    const projStart = new Date(params.projectStart);
    const projEnd   = new Date(params.projectEnd);

    // Snap to the Monday on/before projStart
    const firstDay = new Date(projStart);
    const dow = firstDay.getDay();
    firstDay.setDate(firstDay.getDate() - (dow === 0 ? 6 : dow - 1));

    const weekStarts: Date[] = [];
    const cur = new Date(firstDay);
    while (cur <= projEnd) {
      weekStarts.push(new Date(cur));
      cur.setDate(cur.getDate() + granDays);
    }
    const nCols = weekStarts.length;

    // --- Layout ---
    const MARGIN       = 10;
    const TITLE_H      = params.title ? 28 : 0;
    const TABLE_LEFT   = MARGIN;
    const TABLE_TOP    = MARGIN + TITLE_H + (TITLE_H > 0 ? 4 : 0);
    const ACT_COL_W    = 138;
    const MONTH_ROW_H  = 22;
    const WEEK_ROW_H   = 16;
    const ACT_ROW_H    = 44;
    const maxTimeW     = SLIDE_WIDTH - TABLE_LEFT - ACT_COL_W - MARGIN;
    const TIME_COL_W   = Math.max(28, Math.floor(maxTimeW / nCols));
    const TABLE_W      = ACT_COL_W + nCols * TIME_COL_W;
    const TABLE_H      = MONTH_ROW_H + WEEK_ROW_H + params.rows.length * ACT_ROW_H;

    function dateToX(date: Date): number {
      const totalMs = nCols * granDays * 86400000;
      const fraction = Math.max(0, Math.min(1, (date.getTime() - weekStarts[0].getTime()) / totalMs));
      return TABLE_LEFT + ACT_COL_W + fraction * (nCols * TIME_COL_W);
    }

    // --- Title ---
    if (params.title) {
      const t = slide.shapes.addTextBox(params.title.toUpperCase(), {
        left: TABLE_LEFT, top: MARGIN, width: TABLE_W, height: TITLE_H,
      });
      t.name = `${GANTT_TAG}title`;
      t.fill.clear();
      t.lineFormat.visible = false;
      t.textFrame.textRange.font.bold = true;
      t.textFrame.textRange.font.size = 12;
      t.textFrame.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
    }

    // --- Table ---
    const tableShape = slide.shapes.addTable(2 + params.rows.length, 1 + nCols, {
      left: TABLE_LEFT, top: TABLE_TOP, width: TABLE_W, height: TABLE_H,
    });
    tableShape.name = `${GANTT_TAG}table`;
    const tbl = tableShape.table;

    // Column widths
    tbl.columns.getItemAt(0).width = ACT_COL_W;
    for (let c = 0; c < nCols; c++) tbl.columns.getItemAt(1 + c).width = TIME_COL_W;

    // Row heights
    tbl.rows.getItemAt(0).height = MONTH_ROW_H;
    tbl.rows.getItemAt(1).height = WEEK_ROW_H;
    for (let r = 0; r < params.rows.length; r++) tbl.rows.getItemAt(2 + r).height = ACT_ROW_H;

    // "Activity" header — spans month + week rows
    const actCell = tbl.getCellOrNullObject(0, 0);
    actCell.text = "Activity";
    actCell.resize(2, 1);
    actCell.fill.setSolidColor("#595959");
    actCell.font.color = "#FFFFFF";
    actCell.font.bold = true;
    actCell.font.size = 10;
    actCell.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
    actCell.verticalAlignment   = PowerPoint.TextVerticalAlignment.middle;

    // Month headers (merged cells per month)
    let curMonth = "";
    let mStart = 0;
    type MG = { month: string; start: number; end: number };
    const monthGroups: MG[] = [];
    for (let c = 0; c < nCols; c++) {
      const m = weekStarts[c].toLocaleString("en-US", { month: "long" });
      if (m !== curMonth) {
        if (curMonth) monthGroups.push({ month: curMonth, start: mStart, end: c - 1 });
        curMonth = m; mStart = c;
      }
    }
    monthGroups.push({ month: curMonth, start: mStart, end: nCols - 1 });

    for (const mg of monthGroups) {
      const cell = tbl.getCellOrNullObject(0, 1 + mg.start);
      cell.text = mg.month;
      if (mg.end > mg.start) cell.resize(1, mg.end - mg.start + 1);
      cell.fill.setSolidColor("#595959");
      cell.font.color = "#FFFFFF";
      cell.font.bold  = true;
      cell.font.size  = 10;
      cell.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      cell.verticalAlignment   = PowerPoint.TextVerticalAlignment.middle;
    }

    // Week date headers
    for (let c = 0; c < nCols; c++) {
      const cell = tbl.getCellOrNullObject(1, 1 + c);
      cell.text = String(weekStarts[c].getDate());
      cell.fill.setSolidColor("#747474");
      cell.font.color = "#FFFFFF";
      cell.font.size  = 9;
      cell.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
      cell.verticalAlignment   = PowerPoint.TextVerticalAlignment.middle;
    }

    // Activity rows
    for (let r = 0; r < params.rows.length; r++) {
      const row = params.rows[r];
      const cell = tbl.getCellOrNullObject(2 + r, 0);
      cell.text = row.activity + (row.team ? `\n(${row.team})` : "");
      cell.font.size = 10;
      cell.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      if (r % 2 === 1) {
        for (let c = 0; c < nCols; c++) {
          tbl.getCellOrNullObject(2 + r, 1 + c).fill.setSolidColor("#F5F5F5");
        }
      }
    }

    await context.sync();

    // --- Task bars ---
    const BAR_PAD = 5;
    const BAR_H   = ACT_ROW_H - BAR_PAD * 2;

    for (let rowIdx = 0; rowIdx < params.rows.length; rowIdx++) {
      const row    = params.rows[rowIdx];
      const barTop = TABLE_TOP + MONTH_ROW_H + WEEK_ROW_H + rowIdx * ACT_ROW_H + BAR_PAD;

      for (const task of row.tasks) {
        const x1   = Math.max(TABLE_LEFT + ACT_COL_W + 1, dateToX(new Date(task.start)));
        const x2   = Math.min(TABLE_LEFT + ACT_COL_W + nCols * TIME_COL_W - 1, dateToX(new Date(task.end)));
        const barW = Math.max(22, x2 - x1);

        const bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.homePlate, {
          left: x1, top: barTop, width: barW, height: BAR_H,
        });
        bar.name = `${GANTT_TAG}bar_${rowIdx}`;
        bar.fill.setSolidColor(task.color ?? "#6B9E6B");
        bar.lineFormat.visible = false;
        bar.textFrame.textRange.text = task.label;
        bar.textFrame.textRange.font.color = "#FFFFFF";
        bar.textFrame.textRange.font.size  = 9;
        bar.textFrame.verticalAlignment   = PowerPoint.TextVerticalAlignment.middle;
        bar.textFrame.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.left;
      }

      if (row.note) {
        const lastTask = row.tasks[row.tasks.length - 1];
        if (lastTask) {
          const noteX = dateToX(new Date(lastTask.end)) + 6;
          if (noteX < TABLE_LEFT + TABLE_W - 20) {
            const nb = slide.shapes.addTextBox(row.note, {
              left: noteX, top: barTop,
              width: TABLE_LEFT + TABLE_W - noteX - MARGIN, height: BAR_H,
            });
            nb.name = `${GANTT_TAG}note_${rowIdx}`;
            nb.fill.clear();
            nb.lineFormat.visible = false;
            nb.textFrame.textRange.font.size  = 9;
            nb.textFrame.textRange.font.color = "#C00000";
            nb.textFrame.textRange.font.bold  = true;
            nb.textFrame.verticalAlignment    = PowerPoint.TextVerticalAlignment.middle;
          }
        }
      }
    }

    // --- Today marker ---
    if (params.todayDate) {
      const tx = dateToX(new Date(params.todayDate));
      if (tx > TABLE_LEFT + ACT_COL_W && tx < TABLE_LEFT + TABLE_W) {
        const line = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left: tx - 1, top: TABLE_TOP + MONTH_ROW_H,
          width: 2, height: WEEK_ROW_H + params.rows.length * ACT_ROW_H,
        });
        line.name = `${GANTT_TAG}today_line`;
        line.fill.setSolidColor("#00B4B4");
        line.lineFormat.visible = false;

        const lbl = slide.shapes.addTextBox("Today", {
          left: tx - 18, top: TABLE_TOP + TABLE_H + 2, width: 36, height: 14,
        });
        lbl.name = `${GANTT_TAG}today_label`;
        lbl.fill.clear();
        lbl.lineFormat.visible = false;
        lbl.textFrame.textRange.font.size  = 8;
        lbl.textFrame.textRange.font.color = "#00B4B4";
        lbl.textFrame.textRange.font.bold  = true;
        lbl.textFrame.horizontalAlignment  = PowerPoint.ParagraphHorizontalAlignment.center;
      }
    }

    // --- Data shape (for re-editing) ---
    const dataBox = slide.shapes.addTextBox(JSON.stringify(params), {
      left: 0, top: SLIDE_HEIGHT - 1, width: 20, height: 1,
    });
    dataBox.name = `${GANTT_TAG}data`;
    dataBox.fill.clear();
    dataBox.lineFormat.visible = false;
    dataBox.textFrame.textRange.font.size  = 1;
    dataBox.textFrame.textRange.font.color = "#FFFFFF";

    await context.sync();

    return { type: "success", message: `Built Gantt chart with ${params.rows.length} row(s).` };
  });
}

export type Milestone = {
  label: string;
  date: string; // ISO date "YYYY-MM-DD"
};

export type TimelineParams = {
  milestones: Milestone[];
  startDate: string;
  endDate: string;
};

export async function buildTimeline(params: TimelineParams): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const { milestones, startDate, endDate } = params;

    if (!milestones || milestones.length === 0) {
      return { type: "warning", message: "No milestones provided for the timeline." };
    }

    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const slide = selectedSlides.items[0];

    const startMs = new Date(startDate).getTime();
    const endMs = new Date(endDate).getTime();
    const totalMs = endMs - startMs;

    if (totalMs <= 0) {
      return { type: "warning", message: "End date must be after start date." };
    }

    const LINE_LEFT = 40;
    const LINE_TOP = SLIDE_HEIGHT / 2;
    const LINE_WIDTH = SLIDE_WIDTH - 80;
    const MARKER_SIZE = 10;
    const LABEL_HEIGHT = 24;
    const LABEL_WIDTH = 80;

    // Draw baseline
    const baseline = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: LINE_LEFT,
      top: LINE_TOP - 2,
      width: LINE_WIDTH,
      height: 4,
    });
    baseline.fill.setSolidColor("#323130");
    baseline.lineFormat.visible = false;

    for (const milestone of milestones) {
      const ms = new Date(milestone.date).getTime();
      const x = LINE_LEFT + ((ms - startMs) / totalMs) * LINE_WIDTH;

      // Diamond marker
      const marker = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond, {
        left: x - MARKER_SIZE / 2,
        top: LINE_TOP - MARKER_SIZE / 2,
        width: MARKER_SIZE,
        height: MARKER_SIZE,
      });
      marker.fill.setSolidColor("#0078D4");
      marker.lineFormat.visible = false;

      // Label
      slide.shapes.addTextBox(milestone.label, {
        left: x - LABEL_WIDTH / 2,
        top: LINE_TOP - MARKER_SIZE - LABEL_HEIGHT - 4,
        width: LABEL_WIDTH,
        height: LABEL_HEIGHT,
      });
    }

    await context.sync();

    return {
      type: "success",
      message: `Built timeline with ${milestones.length} milestone(s).`,
    };
  });
}

export type OutlineParams = {
  outline: string;
};

export async function buildSlidesFromOutline(params: OutlineParams): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const { outline } = params;

    if (!outline?.trim()) {
      return { type: "warning", message: "No outline text provided." };
    }

    const lines = outline.split(/\r\n|\r|\n/);
    const slides: { title: string; bullets: string[] }[] = [];
    let current: { title: string; bullets: string[] } | null = null;

    for (const raw of lines) {
      const line = raw.trimEnd();
      if (!line.trim()) continue;

      // H1 (# Title) or unindented text → new slide
      if (/^#\s/.test(line) || (!/^\s/.test(line) && !line.startsWith("-") && !line.startsWith("*"))) {
        current = { title: line.replace(/^#+\s*/, "").trim(), bullets: [] };
        slides.push(current);
      } else if (current) {
        current.bullets.push(line.replace(/^[\s\-*]+/, "").trim());
      }
    }

    if (slides.length === 0) {
      return { type: "warning", message: "Could not parse any slides from the outline." };
    }

    const presentation = context.presentation;

    for (const slideData of slides) {
      const newSlide = presentation.slides.add();
      newSlide.shapes.addTextBox(slideData.title, {
        left: 40,
        top: 40,
        width: SLIDE_WIDTH - 80,
        height: 80,
      });

      if (slideData.bullets.length > 0) {
        newSlide.shapes.addTextBox(slideData.bullets.join("\n"), {
          left: 40,
          top: 140,
          width: SLIDE_WIDTH - 80,
          height: SLIDE_HEIGHT - 180,
        });
      }
    }

    await context.sync();

    return {
      type: "success",
      message: `Created ${slides.length} slide(s) from outline.`,
    };
  });
}

export async function convertTableToGantt(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length !== 1) {
      return { type: "warning", message: "Select exactly one table shape to convert." };
    }

    shapes[0].load("type");
    await context.sync();

    if (shapes[0].type !== "Table") {
      return { type: "warning", message: "Selected shape is not a table." };
    }

    const table = shapes[0].table;
    table.load("rowCount,columnCount");
    await context.sync();

    const rowCount = table.rowCount;
    const colCount = table.columnCount;

    const cellMatrix: PowerPoint.TableCell[][] = [];
    for (let r = 0; r < rowCount; r++) {
      const row: PowerPoint.TableCell[] = [];
      for (let c = 0; c < colCount; c++) {
        const cell = table.getCellOrNullObject(r, c);
        cell.load("text");
        row.push(cell);
      }
      cellMatrix.push(row);
    }
    await context.sync();

    const headers = cellMatrix[0].map((c) => c.text?.toLowerCase()?.trim() ?? "");
    const actCol   = headers.findIndex((h) => h.includes("task") || h.includes("activity") || h.includes("name"));
    const startCol = headers.findIndex((h) => h.includes("start"));
    const endCol   = headers.findIndex((h) => h.includes("end") || h.includes("finish"));

    if (actCol < 0 || startCol < 0 || endCol < 0) {
      return {
        type: "warning",
        message: 'Table must have columns named "activity"/"task", "start", and "end"/"finish".',
      };
    }

    const rows: GanttRowEntry[] = [];
    let projectStart = "";
    let projectEnd   = "";

    for (let r = 1; r < rowCount; r++) {
      const activity = cellMatrix[r][actCol].text?.trim() ?? "";
      const start    = cellMatrix[r][startCol].text?.trim() ?? "";
      const end      = cellMatrix[r][endCol].text?.trim() ?? "";
      if (!activity || !start || !end) continue;
      if (!projectStart || start < projectStart) projectStart = start;
      if (!projectEnd   || end   > projectEnd)   projectEnd   = end;
      rows.push({ activity, tasks: [{ label: activity, start, end }] });
    }

    if (rows.length === 0) {
      return { type: "warning", message: "No valid rows found in the table." };
    }

    return buildGanttChart({ projectStart, projectEnd, rows });
  });
}

export async function pasteAsGrid(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    // Read clipboard text
    let clipboardText = "";
    try {
      clipboardText = await navigator.clipboard.readText();
    } catch {
      return {
        type: "warning",
        message: "Clipboard access denied. Please allow clipboard permission and try again.",
      };
    }

    const rows = clipboardText
      .split(/\r\n|\r|\n/)
      .map((r) => r.trim())
      .filter(Boolean);

    if (rows.length === 0) {
      return { type: "info", message: "Clipboard is empty or contains no text rows." };
    }

    const cols = Math.ceil(Math.sqrt(rows.length));
    const gridRows = Math.ceil(rows.length / cols);

    const gap = 12;
    const cellWidth = (SLIDE_WIDTH - 40 - (cols - 1) * gap) / cols;
    const cellHeight = Math.min(80, (SLIDE_HEIGHT - 40 - (gridRows - 1) * gap) / gridRows);

    const slide = selectedSlides.items[0];

    rows.forEach((text, idx) => {
      const col = idx % cols;
      const row = Math.floor(idx / cols);
      slide.shapes.addTextBox(text, {
        left: 20 + col * (cellWidth + gap),
        top: 20 + row * (cellHeight + gap),
        width: cellWidth,
        height: cellHeight,
      });
    });

    await context.sync();

    return {
      type: "success",
      message: `Created ${rows.length} card(s) in a ${gridRows}×${cols} grid.`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Slide organisation
// ─────────────────────────────────────────────────────────────────

const UNUSED_DIVIDER_TAG = "JOLIFY_UNUSED_DIVIDER";

export async function moveToUnusedSection(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    const allSlides = context.presentation.slides;
    selectedSlides.load("items");
    allSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "warning", message: "Select at least one slide in the slide panel first." };
    }

    // Load IDs so we can compare slides
    allSlides.items.forEach((s) => s.load("id"));
    selectedSlides.items.forEach((s) => s.load("id"));

    // Check every slide for the divider tag
    const tagChecks = allSlides.items.map((s) =>
      s.tags.getItemOrNullObject(UNUSED_DIVIDER_TAG)
    );
    await context.sync();

    const dividerIdx  = tagChecks.findIndex((t) => !t.isNullObject);
    const dividerSlide = dividerIdx >= 0 ? allSlides.items[dividerIdx] : null;
    const selectedIds  = new Set(selectedSlides.items.map((s) => s.id));

    // Filter out the divider itself if the user somehow selected it
    const slidesToMove = selectedSlides.items.filter(
      (s) => !dividerSlide || s.id !== dividerSlide.id
    );

    if (slidesToMove.length === 0) {
      return { type: "info", message: "Nothing to move (divider slide is excluded)." };
    }

    // Guard: don't let the user move every real slide to unused
    const totalReal = allSlides.items.length - (dividerSlide ? 1 : 0);
    if (slidesToMove.length >= totalReal) {
      return { type: "warning", message: "Cannot move all slides to Unused Slides." };
    }

    // Create divider slide if one doesn't exist yet
    let totalSlides = allSlides.items.length;
    if (!dividerSlide) {
      allSlides.add();
      await context.sync();
      allSlides.load("items");
      await context.sync();
      totalSlides = allSlides.items.length;

      const divSlide = allSlides.items[totalSlides - 1];
      divSlide.tags.add(UNUSED_DIVIDER_TAG, "true");

      const tb = divSlide.shapes.addTextBox("── Unused Slides ──", {
        left: 40,
        top: SLIDE_HEIGHT / 2 - 40,
        width: SLIDE_WIDTH - 80,
        height: 60,
      });
      tb.textFrame.textRange.font.size  = 24;
      tb.textFrame.textRange.font.color = "#A0A0A0";
      tb.textFrame.horizontalAlignment  = PowerPoint.ParagraphHorizontalAlignment.center;
      await context.sync();
    }

    // Move each selected slide to the last position.
    // Because moveTo(last) naturally pushes the divider one spot left each time,
    // the divider always ends up just before all the moved slides.
    for (const slide of slidesToMove) {
      slide.moveTo(totalSlides - 1);
      await context.sync();
    }

    const n = slidesToMove.length;
    return {
      type: "success",
      message: `Moved ${n} slide${n !== 1 ? "s" : ""} to "Unused Slides".`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 3 — Dialog wrappers
// ─────────────────────────────────────────────────────────────────

export async function openGanttDialog(): Promise<ActionResult> {
  // Check current slide for an existing Gantt data shape to pre-populate the dialog
  let existingDataParam = "";
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();
      if (slides.items.length === 0) return;
      const shapeColl = slides.items[0].shapes;
      shapeColl.load("items");
      await context.sync();
      shapeColl.items.forEach((s) => s.load("name"));
      await context.sync();
      const dataShape = shapeColl.items.find((s) => s.name === `${GANTT_TAG}data`);
      if (dataShape) {
        dataShape.textFrame.textRange.load("text");
        await context.sync();
        const json = dataShape.textFrame.textRange.text?.trim();
        if (json) existingDataParam = btoa(unescape(encodeURIComponent(json)));
      }
    });
  } catch { /* no existing data */ }

  const url = existingDataParam
    ? `dialogs/gantt-builder.html?data=${encodeURIComponent(existingDataParam)}`
    : "dialogs/gantt-builder.html";

  const params = await openDialog<GanttChartParams>(url);
  if (!params) return { type: "info", message: "Gantt chart creation cancelled." };
  return buildGanttChart(params);
}

export async function openTimelineDialog(): Promise<ActionResult> {
  const params = await openDialog<TimelineParams>("dialogs/timeline-builder.html");
  if (!params) return { type: "info", message: "Timeline creation cancelled." };
  return buildTimeline(params);
}

export async function openSlideOutlineDialog(): Promise<ActionResult> {
  const params = await openDialog<OutlineParams>("dialogs/slide-outline.html");
  if (!params) return { type: "info", message: "Slide creation cancelled." };
  return buildSlidesFromOutline(params);
}
