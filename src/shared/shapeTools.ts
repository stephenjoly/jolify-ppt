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

type SavedFillStyle =
  | { type: "NoFill" }
  | { type: "Solid"; foregroundColor: string; transparency: number };

type SavedOutlineStyle = {
  color: string;
  dashStyle: string;
  style: string;
  transparency: number;
  visible: boolean;
  weight: number;
};

type SavedTextStyle = {
  bold: boolean | null;
  color: string;
  italic: boolean | null;
  size: number | null;
  underline: string;
};

type SavedShapeStyle = {
  fill: SavedFillStyle;
  outline: SavedOutlineStyle;
  text: SavedTextStyle;
};

let savedPosition: ShapePosition | null = null;
let savedSize: ShapeSize | null = null;
let savedFillStyle: SavedFillStyle | null = null;
let savedOutlineStyle: SavedOutlineStyle | null = null;
let savedShapeStyle: SavedShapeStyle | null = null;
const SAVED_POSITION_KEY = "__jolify_saved_position__";
const SAVED_SIZE_KEY = "__jolify_saved_size__";
const SAVED_FILL_STYLE_KEY = "__jolify_saved_fill_style__";
const SAVED_OUTLINE_STYLE_KEY = "__jolify_saved_outline_style__";
const SAVED_SHAPE_STYLE_KEY = "__jolify_saved_shape_style__";

function getDocumentSettings(): Office.Settings | null {
  return Office.context?.document?.settings ?? null;
}

function parseStoredValue<T>(value: unknown): T | null {
  if (value == null) {
    return null;
  }

  if (typeof value === "string") {
    try {
      return JSON.parse(value) as T;
    } catch {
      return null;
    }
  }

  return value as T;
}

function refreshDocumentSettings(settings: Office.Settings): Promise<void> {
  return new Promise((resolve, reject) => {
    settings.refreshAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(new Error(result.error?.message || "Could not refresh saved Jolify settings."));
    });
  });
}

function saveDocumentSettings(settings: Office.Settings): Promise<void> {
  return new Promise((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(new Error(result.error?.message || "Could not save Jolify settings."));
    });
  });
}

async function loadSavedGeometry(): Promise<void> {
  const settings = getDocumentSettings();
  if (!settings) {
    return;
  }

  await refreshDocumentSettings(settings);
  savedPosition = parseStoredValue<ShapePosition>(settings.get(SAVED_POSITION_KEY));
  savedSize = parseStoredValue<ShapeSize>(settings.get(SAVED_SIZE_KEY));
  savedFillStyle = parseStoredValue<SavedFillStyle>(settings.get(SAVED_FILL_STYLE_KEY));
  savedOutlineStyle = parseStoredValue<SavedOutlineStyle>(settings.get(SAVED_OUTLINE_STYLE_KEY));
  savedShapeStyle = parseStoredValue<SavedShapeStyle>(settings.get(SAVED_SHAPE_STYLE_KEY));
}

async function persistSavedGeometry(): Promise<void> {
  const settings = getDocumentSettings();
  if (!settings) {
    return;
  }

  if (savedPosition) {
    settings.set(SAVED_POSITION_KEY, savedPosition);
  } else {
    settings.remove(SAVED_POSITION_KEY);
  }

  if (savedSize) {
    settings.set(SAVED_SIZE_KEY, savedSize);
  } else {
    settings.remove(SAVED_SIZE_KEY);
  }

  if (savedFillStyle) {
    settings.set(SAVED_FILL_STYLE_KEY, savedFillStyle);
  } else {
    settings.remove(SAVED_FILL_STYLE_KEY);
  }

  if (savedOutlineStyle) {
    settings.set(SAVED_OUTLINE_STYLE_KEY, savedOutlineStyle);
  } else {
    settings.remove(SAVED_OUTLINE_STYLE_KEY);
  }

  if (savedShapeStyle) {
    settings.set(SAVED_SHAPE_STYLE_KEY, savedShapeStyle);
  } else {
    settings.remove(SAVED_SHAPE_STYLE_KEY);
  }

  await saveDocumentSettings(settings);
}

async function getSelectedShapes(context: PowerPoint.RequestContext) {
  const selection = context.presentation.getSelectedShapes();
  selection.load("items");
  await context.sync();

  if (selection.items.length > 0) {
    return selection.items;
  }

  const textRange = context.presentation.getSelectedTextRangeOrNullObject();
  textRange.load("isNullObject");
  await context.sync();

  if (textRange.isNullObject) {
    return [];
  }

  const parentTextFrame = textRange.getParentTextFrame();
  const parentShape = parentTextFrame.getParentShape();
  parentShape.load("id");
  await context.sync();

  return [parentShape];
}

type SelectedTextShapesResult = {
  shapes: PowerPoint.Shape[];
  skippedCount: number;
};

async function getSelectedTextShapes(
  context: PowerPoint.RequestContext,
  emptyMessage: string,
): Promise<SelectedTextShapesResult | ActionResult> {
  const shapes = await getSelectedShapes(context);
  if (shapes.length < 1) {
    return {
      type: "warning",
      message: emptyMessage,
    };
  }

  shapes.forEach((shape) => {
    shape.load("type,hasText");
  });
  await context.sync();

  const textShapes = shapes.filter((shape) => shape.type === "TextBox" || shape.hasText);
  if (textShapes.length < 1) {
    return {
      type: "warning",
      message: emptyMessage,
    };
  }

  return {
    shapes: textShapes,
    skippedCount: shapes.length - textShapes.length,
  };
}

function getIgnoredShapeNote(skippedCount: number): string {
  return skippedCount > 0
    ? ` Ignored ${skippedCount} non-text shape${skippedCount !== 1 ? "s" : ""}.`
    : "";
}

export async function copyPositionOnly(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape or click into table/text content to copy its position.",
      };
    }

    const shape = shapes[0];
    shape.load("left,top");
    await context.sync();

    savedPosition = {
      left: shape.left,
      top: shape.top,
    };
    await persistSavedGeometry();

    return {
      type: "success",
      message: "Saved position from the first selected shape.",
    };
  });
}

export async function copySizeOnly(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape or click into table/text content to copy its size.",
      };
    }

    const shape = shapes[0];
    shape.load("width,height");
    await context.sync();

    savedSize = {
      width: shape.width,
      height: shape.height,
    };
    await persistSavedGeometry();

    return {
      type: "success",
      message: "Saved size from the first selected shape.",
    };
  });
}

export async function copyPositionAndSize(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape or click into table/text content to copy its position & size.",
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
    await persistSavedGeometry();

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
  await loadSavedGeometry();
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
        message: "Select one or more shapes, or click into table/text content, to paste the saved position.",
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
  await loadSavedGeometry();
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
        message: "Select one or more shapes, or click into table/text content, to paste the saved size.",
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
  await loadSavedGeometry();
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
        message: "Select one or more shapes, or click into table/text content, to paste the saved position & size.",
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
const CENTER_STICKER_NAME = "__jolify_center_sticker__";
const CENTER_STICKER_WIDTH = 148.8189;
const CENTER_STICKER_HEIGHT = 99.2126;
const CENTER_STICKER_TEXT = "Placeholder text";
const CENTER_STICKER_FILL_COLOR = "#D9D9D9";
const CENTER_STICKER_OUTLINE_COLOR = "#BFBFBF";
const CENTER_STICKER_CASCADE_OFFSET = 10;
const POSITION_TOLERANCE = 0.5;

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

export async function toggleDraftSticker(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (slides.items.length === 0) {
      return { type: "warning", message: "No slides found." };
    }

    let hasAnySticker = false;

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();
      shapes.items.forEach((s) => s.load("name"));
      await context.sync();

      if (shapes.items.some((s) => s.name === DRAFT_STICKER_NAME)) {
        hasAnySticker = true;
        break;
      }
    }

    return hasAnySticker ? removeDraftSticker() : addDraftSticker();
  });
}

const SLIDE_WIDTH = 960;  // points, widescreen default (same assumption as addDraftSticker)
const SLIDE_HEIGHT = 540; // points, widescreen default

type AlignType = "left" | "centerH" | "right" | "top" | "middleV" | "bottom" | "distributeH" | "distributeV";
type StretchEdge = "left" | "right" | "top" | "bottom";

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

async function stretchShapesToReferenceEdge(edge: StretchEdge): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes. The first selected shape is the reference.",
      };
    }

    shapes.forEach((shape) => {
      shape.load("left,top,width,height");
    });
    await context.sync();

    const [referenceShape, ...targetShapes] = shapes;
    const referenceRight = referenceShape.left + referenceShape.width;
    const referenceBottom = referenceShape.top + referenceShape.height;

    let updatedCount = 0;
    let skippedCount = 0;

    targetShapes.forEach((shape) => {
      switch (edge) {
        case "left": {
          const fixedRight = shape.left + shape.width;
          const newWidth = fixedRight - referenceShape.left;
          if (newWidth <= 0) {
            skippedCount += 1;
            return;
          }
          shape.left = referenceShape.left;
          shape.width = newWidth;
          updatedCount += 1;
          return;
        }
        case "right": {
          const newWidth = referenceRight - shape.left;
          if (newWidth <= 0) {
            skippedCount += 1;
            return;
          }
          shape.width = newWidth;
          updatedCount += 1;
          return;
        }
        case "top": {
          const fixedBottom = shape.top + shape.height;
          const newHeight = fixedBottom - referenceShape.top;
          if (newHeight <= 0) {
            skippedCount += 1;
            return;
          }
          shape.top = referenceShape.top;
          shape.height = newHeight;
          updatedCount += 1;
          return;
        }
        case "bottom": {
          const newHeight = referenceBottom - shape.top;
          if (newHeight <= 0) {
            skippedCount += 1;
            return;
          }
          shape.height = newHeight;
          updatedCount += 1;
        }
      }
    });

    if (updatedCount === 0) {
      return {
        type: "warning",
        message: "No target shapes could be resized to that reference edge without collapsing.",
      };
    }

    await context.sync();

    const edgeLabel = edge.charAt(0).toUpperCase() + edge.slice(1);
    const skippedNote =
      skippedCount > 0
        ? ` Skipped ${skippedCount} shape${skippedCount !== 1 ? "s" : ""} that would have collapsed.`
        : "";

    return {
      type: "success",
      message: `Stretched ${updatedCount} shape${updatedCount !== 1 ? "s" : ""} to match the reference ${edgeLabel.toLowerCase()} edge.${skippedNote}`,
    };
  });
}

function getSavedFillFromShape(shape: PowerPoint.Shape): SavedFillStyle | null {
  if (shape.fill.type === "NoFill") {
    return { type: "NoFill" };
  }

  if (shape.fill.type === "Solid") {
    return {
      type: "Solid",
      foregroundColor: shape.fill.foregroundColor,
      transparency: shape.fill.transparency,
    };
  }

  return null;
}

function applySavedFillStyle(shape: PowerPoint.Shape, style: SavedFillStyle): void {
  if (style.type === "NoFill") {
    shape.fill.clear();
    return;
  }

  shape.fill.setSolidColor(style.foregroundColor);
  shape.fill.transparency = style.transparency;
}

function getSavedOutlineFromShape(shape: PowerPoint.Shape): SavedOutlineStyle {
  return {
    color: shape.lineFormat.color,
    dashStyle: String(shape.lineFormat.dashStyle),
    style: String(shape.lineFormat.style),
    transparency: shape.lineFormat.transparency,
    visible: shape.lineFormat.visible,
    weight: shape.lineFormat.weight,
  };
}

function applySavedOutlineStyle(shape: PowerPoint.Shape, style: SavedOutlineStyle): void {
  shape.lineFormat.visible = style.visible;
  shape.lineFormat.color = style.color;
  shape.lineFormat.dashStyle = style.dashStyle as any;
  shape.lineFormat.style = style.style as any;
  shape.lineFormat.transparency = style.transparency;
  shape.lineFormat.weight = style.weight;
}

function getSavedTextStyleFromShape(shape: PowerPoint.Shape): SavedTextStyle {
  return {
    bold: shape.textFrame.textRange.font.bold,
    color: shape.textFrame.textRange.font.color,
    italic: shape.textFrame.textRange.font.italic,
    size: shape.textFrame.textRange.font.size ?? null,
    underline: String(shape.textFrame.textRange.font.underline),
  };
}

function applySavedTextStyle(shape: PowerPoint.Shape, style: SavedTextStyle): void {
  shape.textFrame.textRange.font.bold = style.bold;
  shape.textFrame.textRange.font.color = style.color;
  shape.textFrame.textRange.font.italic = style.italic;
  if (style.size) {
    shape.textFrame.textRange.font.size = style.size;
  }
  shape.textFrame.textRange.font.underline = style.underline as any;
}

function getSavedShapeStyleFromShape(shape: PowerPoint.Shape): SavedShapeStyle | null {
  const fill = getSavedFillFromShape(shape);
  if (!fill) {
    return null;
  }

  return {
    fill,
    outline: getSavedOutlineFromShape(shape),
    text: getSavedTextStyleFromShape(shape),
  };
}

function applySavedShapeStyle(shape: PowerPoint.Shape, style: SavedShapeStyle): void {
  applySavedFillStyle(shape, style.fill);
  applySavedOutlineStyle(shape, style.outline);
  applySavedTextStyle(shape, style.text);
}

export async function matchFillStyle(): Promise<ActionResult> {
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

    const savedStyle = getSavedFillFromShape(source);
    if (!savedStyle) {
      return {
        type: "warning",
        message: "Only solid color and no-fill can be matched. Gradient, pattern, and picture fills aren't supported by the Office.js API.",
      };
    }

    shapes.slice(1).forEach((shape) => {
      applySavedFillStyle(shape, savedStyle);
    });

    await context.sync();

    return {
      type: "success",
      message: `Matched fill on ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function copyFillToClipboard(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its fill.",
      };
    }

    const source = shapes[0];
    source.fill.load("type,foregroundColor,transparency");
    await context.sync();

    const savedStyle = getSavedFillFromShape(source);
    if (!savedStyle) {
      return {
        type: "warning",
        message: "Only solid color and no-fill can be copied. Gradient, pattern, and picture fills aren't supported by the Office.js API.",
      };
    }

    savedFillStyle = savedStyle;
    await persistSavedGeometry();

    return {
      type: "success",
      message: "Saved fill from the first selected shape.",
    };
  });
}

export async function pasteFillFromClipboard(): Promise<ActionResult> {
  await loadSavedGeometry();
  if (!savedFillStyle) {
    return {
      type: "warning",
      message: "Copy a fill first.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to paste fill onto.",
      };
    }

    shapes.forEach((shape) => {
      applySavedFillStyle(shape, savedFillStyle!);
    });
    await context.sync();

    return {
      type: "success",
      message: `Pasted fill to ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
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

export async function matchOutlineStyle(): Promise<ActionResult> {
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

    const savedStyle = getSavedOutlineFromShape(source);

    shapes.slice(1).forEach((shape) => {
      applySavedOutlineStyle(shape, savedStyle);
    });

    await context.sync();

    return {
      type: "success",
      message: `Matched outline on ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function copyOutlineToClipboard(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its outline.",
      };
    }

    const source = shapes[0];
    source.lineFormat.load("color,dashStyle,style,transparency,visible,weight");
    await context.sync();

    savedOutlineStyle = getSavedOutlineFromShape(source);
    await persistSavedGeometry();

    return {
      type: "success",
      message: "Saved outline from the first selected shape.",
    };
  });
}

export async function pasteOutlineFromClipboard(): Promise<ActionResult> {
  await loadSavedGeometry();
  if (!savedOutlineStyle) {
    return {
      type: "warning",
      message: "Copy an outline first.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to paste outline onto.",
      };
    }

    shapes.forEach((shape) => {
      applySavedOutlineStyle(shape, savedOutlineStyle!);
    });
    await context.sync();

    return {
      type: "success",
      message: `Pasted outline to ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function clearFill(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to remove its fill.",
      };
    }

    shapes.forEach((shape) => {
      shape.fill.clear();
    });

    await context.sync();

    return {
      type: "success",
      message: `Removed fill from ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function applyFillColor(color: string): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select one or more shapes before applying a fill color.",
      };
    }

    shapes.forEach((shape) => {
      shape.fill.setSolidColor(color);
    });

    await context.sync();

    return {
      type: "success",
      message: `Applied fill color to ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function clearOutline(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);

    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to remove its outline.",
      };
    }

    shapes.forEach((shape) => {
      shape.lineFormat.visible = false;
    });

    await context.sync();

    return {
      type: "success",
      message: `Removed outline from ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function applyOutlineColor(color: string): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select one or more shapes before applying an outline color.",
      };
    }

    shapes.forEach((shape) => {
      shape.lineFormat.visible = true;
      shape.lineFormat.color = color;
    });

    await context.sync();

    return {
      type: "success",
      message: `Applied outline color to ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

export async function applyFontColor(color: string): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selectedTextRange = context.presentation.getSelectedTextRangeOrNullObject();
    selectedTextRange.load("isNullObject");
    await context.sync();

    if (!selectedTextRange.isNullObject) {
      selectedTextRange.font.color = color;
      await context.sync();

      return {
        type: "success",
        message: "Applied font color to the selected text.",
      };
    }

    const selected = await getSelectedTextShapes(
      context,
      "Select one or more text shapes before applying a font color.",
    );

    if ("type" in selected) {
      return selected;
    }

    selected.shapes.forEach((shape) => {
      shape.textFrame.textRange.font.color = color;
    });

    await context.sync();

    return {
      type: "success",
      message: `Applied font color to ${selected.shapes.length} text shape${selected.shapes.length !== 1 ? "s" : ""}.${getIgnoredShapeNote(selected.skippedCount)}`,
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
export const stretchToLeftEdge = (): Promise<ActionResult> => stretchShapesToReferenceEdge("left");
export const stretchToRightEdge = (): Promise<ActionResult> => stretchShapesToReferenceEdge("right");
export const stretchToTopEdge = (): Promise<ActionResult> => stretchShapesToReferenceEdge("top");
export const stretchToBottomEdge = (): Promise<ActionResult> => stretchShapesToReferenceEdge("bottom");

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

export async function removeTextMargins(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selected = await getSelectedTextShapes(
      context,
      "Select at least one text box to remove its text margins.",
    );
    if ("type" in selected) {
      return selected;
    }

    selected.shapes.forEach((shape) => {
      shape.textFrame.leftMargin = 0;
      shape.textFrame.rightMargin = 0;
      shape.textFrame.topMargin = 0;
      shape.textFrame.bottomMargin = 0;
    });
    await context.sync();

    return {
      type: "success",
      message:
        `Removed text margins from ${selected.shapes.length} shape${selected.shapes.length !== 1 ? "s" : ""}.` +
        getIgnoredShapeNote(selected.skippedCount),
    };
  });
}

export async function disableTextAutofit(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selected = await getSelectedTextShapes(
      context,
      "Select at least one text box to turn off autofit.",
    );
    if ("type" in selected) {
      return selected;
    }

    selected.shapes.forEach((shape) => {
      shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    });
    await context.sync();

    return {
      type: "success",
      message:
        `Turned off autofit for ${selected.shapes.length} shape${selected.shapes.length !== 1 ? "s" : ""}.` +
        getIgnoredShapeNote(selected.skippedCount),
    };
  });
}

export async function createCenterSticker(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    const slide = selectedSlides.items[0];
    const targetLeft = (SLIDE_WIDTH - CENTER_STICKER_WIDTH) / 2;
    const targetTop = (SLIDE_HEIGHT - CENTER_STICKER_HEIGHT) / 2;
    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();

    shapes.items.forEach((shape) => {
      shape.load("name,left,top,width,height,hasText");
      shape.textFrame.load("autoSizeSetting,leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment");
      shape.textFrame.textRange.load("text");
      shape.textFrame.textRange.paragraphFormat.load("horizontalAlignment");
      shape.fill.load("type,foregroundColor");
      shape.lineFormat.load("visible,color,weight");
    });
    await context.sync();

    const existingStickerPositions = new Set<number>();
    shapes.items.forEach((shape) => {
      const usesCenterStickerName = shape.name === CENTER_STICKER_NAME;
      const looksLikeLegacyCenterSticker =
        shape.width === CENTER_STICKER_WIDTH &&
        shape.height === CENTER_STICKER_HEIGHT &&
        shape.hasText &&
        shape.textFrame.textRange.text === CENTER_STICKER_TEXT &&
        shape.textFrame.autoSizeSetting === PowerPoint.ShapeAutoSize.autoSizeNone &&
        shape.textFrame.leftMargin === 0 &&
        shape.textFrame.rightMargin === 0 &&
        shape.textFrame.topMargin === 0 &&
        shape.textFrame.bottomMargin === 0 &&
        shape.textFrame.verticalAlignment === PowerPoint.TextVerticalAlignment.middle &&
        shape.textFrame.textRange.paragraphFormat.horizontalAlignment === PowerPoint.ParagraphHorizontalAlignment.center &&
        shape.fill.type === "Solid" &&
        shape.fill.foregroundColor === CENTER_STICKER_FILL_COLOR &&
        shape.lineFormat.visible &&
        shape.lineFormat.color === CENTER_STICKER_OUTLINE_COLOR &&
        Math.abs(shape.lineFormat.weight - 1) <= POSITION_TOLERANCE;

      if (!usesCenterStickerName && !looksLikeLegacyCenterSticker) {
        return;
      }

      const leftStep = (shape.left - targetLeft) / CENTER_STICKER_CASCADE_OFFSET;
      const topStep = (shape.top - targetTop) / CENTER_STICKER_CASCADE_OFFSET;
      const roundedLeftStep = Math.round(leftStep);
      const roundedTopStep = Math.round(topStep);

      if (
        Math.abs(leftStep - roundedLeftStep) <= POSITION_TOLERANCE &&
        Math.abs(topStep - roundedTopStep) <= POSITION_TOLERANCE &&
        roundedLeftStep === roundedTopStep &&
        roundedLeftStep >= 0
      ) {
        existingStickerPositions.add(roundedLeftStep);
      }
    });

    let cascadeIndex = 0;
    while (existingStickerPositions.has(cascadeIndex)) {
      cascadeIndex += 1;
    }

    const sticker = slide.shapes.addTextBox(CENTER_STICKER_TEXT, {
      left: targetLeft + cascadeIndex * CENTER_STICKER_CASCADE_OFFSET,
      top: targetTop + cascadeIndex * CENTER_STICKER_CASCADE_OFFSET,
      width: CENTER_STICKER_WIDTH,
      height: CENTER_STICKER_HEIGHT,
    });

    sticker.name = CENTER_STICKER_NAME;
    sticker.fill.setSolidColor(CENTER_STICKER_FILL_COLOR);
    sticker.lineFormat.visible = true;
    sticker.lineFormat.color = CENTER_STICKER_OUTLINE_COLOR;
    sticker.lineFormat.weight = 1;
    sticker.textFrame.leftMargin = 0;
    sticker.textFrame.rightMargin = 0;
    sticker.textFrame.topMargin = 0;
    sticker.textFrame.bottomMargin = 0;
    sticker.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
    sticker.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    sticker.textFrame.wordWrap = true;
    sticker.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
    sticker.load("id");
    await context.sync();

    slide.setSelectedShapes([sticker.id]);
    await context.sync();

    return {
      type: "success",
      message:
        cascadeIndex === 0
          ? "Created a centered sticker."
          : `Created a centered sticker with a ${cascadeIndex * CENTER_STICKER_CASCADE_OFFSET}pt cascade offset.`,
    };
  });
}

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
        distributeShapesOnAxis(shapes, "horizontal");
        break;
      }
      case "distributeV": {
        distributeShapesOnAxis(shapes, "vertical");
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

function distributeShapesOnAxis(
  shapes: PowerPoint.Shape[],
  axis: "horizontal" | "vertical",
): void {
  const sorted = [...shapes].sort((a, b) =>
    axis === "horizontal" ? a.left - b.left : a.top - b.top,
  );

  if (axis === "horizontal") {
    const span = sorted[sorted.length - 1].left + sorted[sorted.length - 1].width - sorted[0].left;
    const totalWidth = sorted.reduce((sum, shape) => sum + shape.width, 0);
    const gap = (span - totalWidth) / (sorted.length - 1);
    let cursor = sorted[0].left;
    sorted.forEach((shape) => {
      shape.left = cursor;
      cursor += shape.width + gap;
    });
    return;
  }

  const span = sorted[sorted.length - 1].top + sorted[sorted.length - 1].height - sorted[0].top;
  const totalHeight = sorted.reduce((sum, shape) => sum + shape.height, 0);
  const gap = (span - totalHeight) / (sorted.length - 1);
  let cursor = sorted[0].top;
  sorted.forEach((shape) => {
    shape.top = cursor;
    cursor += shape.height + gap;
  });
}

export async function distributeHandVAndGroup(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 3) {
      return { type: "warning", message: "Select at least 3 shapes to distribute on both axes and group." };
    }

    shapes.forEach((s) => s.load("left,top,width,height"));
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "error", message: "Could not determine the current slide." };
    }

    distributeShapesOnAxis(shapes, "horizontal");
    distributeShapesOnAxis(shapes, "vertical");

    await context.sync();

    selectedSlides.items[0].shapes.addGroup(shapes);
    await context.sync();

    return { type: "success", message: `Distributed horizontally and vertically, then grouped ${shapes.length} shapes.` };
  });
}

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

export async function matchShapeStyle(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 2) {
      return {
        type: "warning",
        message: "Select at least 2 shapes — style is matched from the first to the rest.",
      };
    }

    const source = shapes[0];
    source.fill.load("type,foregroundColor,transparency");
    source.lineFormat.load("color,dashStyle,style,transparency,visible,weight");
    source.textFrame.textRange.font.load("bold,color,italic,size,underline");
    await context.sync();

    const savedStyle = getSavedShapeStyleFromShape(source);
    if (!savedStyle) {
      return {
        type: "warning",
        message: "Only solid color and no-fill can be matched as full style. Gradient, pattern, and picture fills aren't supported by the Office.js API.",
      };
    }

    shapes.slice(1).forEach((shape) => {
      applySavedShapeStyle(shape, savedStyle);
    });

    await context.sync();

    return {
      type: "success",
      message: `Matched style on ${shapes.length - 1} shape${shapes.length - 1 !== 1 ? "s" : ""}.`,
    };
  });
}

export async function copyStyleToClipboard(): Promise<ActionResult> {
  return PowerPoint.run(async (context) => {
    await loadSavedGeometry();
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to copy its style.",
      };
    }

    const source = shapes[0];
    source.fill.load("type,foregroundColor,transparency");
    source.lineFormat.load("color,dashStyle,style,transparency,visible,weight");
    source.textFrame.textRange.font.load("bold,color,italic,size,underline");
    await context.sync();

    const savedStyle = getSavedShapeStyleFromShape(source);
    if (!savedStyle) {
      return {
        type: "warning",
        message: "Only solid color and no-fill can be copied as full style. Gradient, pattern, and picture fills aren't supported by the Office.js API.",
      };
    }

    savedShapeStyle = savedStyle;
    await persistSavedGeometry();

    return {
      type: "success",
      message: "Saved style from the first selected shape.",
    };
  });
}

export async function pasteStyleFromClipboard(): Promise<ActionResult> {
  await loadSavedGeometry();
  if (!savedShapeStyle) {
    return {
      type: "warning",
      message: "Copy a style first.",
    };
  }

  return PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapes(context);
    if (shapes.length < 1) {
      return {
        type: "warning",
        message: "Select at least one shape to paste style onto.",
      };
    }

    shapes.forEach((shape) => {
      applySavedShapeStyle(shape, savedShapeStyle!);
    });
    await context.sync();

    return {
      type: "success",
      message: `Pasted style to ${shapes.length} shape${shapes.length !== 1 ? "s" : ""}.`,
    };
  });
}

// ─────────────────────────────────────────────────────────────────
// Phase 2 — Dialog-based helpers
// ─────────────────────────────────────────────────────────────────

export function openDialog<T>(
  relativeUrl: string,
  dialogOptions?: { height: number; width: number },
): Promise<T | null> {
  return new Promise((resolve) => {
    const separator = relativeUrl.includes("?") ? "&" : "?";
    const url = `${window.location.origin}${window.location.pathname.replace(/\/[^/]+$/, "/")}${relativeUrl}${separator}v=${Date.now()}`;
    let dialog: Office.Dialog;

    Office.context.ui.displayDialogAsync(
      url,
      {
        height: dialogOptions?.height ?? 50,
        width: dialogOptions?.width ?? 40,
        displayInIframe: true,
      },
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

// ─────────────────────────────────────────────────────────────────
// Phase 2 — Dialog wrappers (open dialog then execute)
// ─────────────────────────────────────────────────────────────────

export async function openGridDialog(): Promise<ActionResult> {
  const params = await openDialog<GridParams>("dialogs/grid-builder.html");
  if (!params) return { type: "info", message: "Grid creation cancelled." };
  return createGrid(params);
}

export type WeekdayRangeParams = {
  startDate: string;
  endDate: string;
  weekday: number;
};

type WeekdayRangeResult = {
  type: ActionResult["type"];
  message: string;
  output: string;
};

const WEEKDAY_NAMES = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
] as const;

function parseLocalDate(value: string): Date | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return null;
  }

  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  date.setHours(0, 0, 0, 0);
  return date;
}

function formatLocalDate(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export function buildWeekdayRangeResult(params: WeekdayRangeParams): WeekdayRangeResult {
  const { startDate, endDate, weekday } = params;
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);

  if (!start || !end) {
    return {
      type: "warning",
      message: "Enter a valid start date and end date.",
      output: "",
    };
  }

  if (!Number.isInteger(weekday) || weekday < 0 || weekday > 6) {
    return {
      type: "warning",
      message: "Choose a valid weekday.",
      output: "",
    };
  }

  if (start.getTime() > end.getTime()) {
    return {
      type: "warning",
      message: "Start date must be on or before end date.",
      output: "",
    };
  }

  const dates: string[] = [];
  const cursor = new Date(start);
  const offset = (weekday - cursor.getDay() + 7) % 7;
  cursor.setDate(cursor.getDate() + offset);

  while (cursor.getTime() <= end.getTime()) {
    dates.push(formatLocalDate(cursor));
    cursor.setDate(cursor.getDate() + 7);
  }

  const weekdayName = WEEKDAY_NAMES[weekday];
  if (dates.length === 0) {
    return {
      type: "info",
      message: `No ${weekdayName}s found in that range.`,
      output: "",
    };
  }

  return {
    type: "success",
    message: `Found ${dates.length} ${weekdayName}${dates.length !== 1 ? "s" : ""} between ${startDate} and ${endDate}.`,
    output: dates.join("\n"),
  };
}

export function openWeekdayRangeDialog(): Promise<ActionResult> {
  return new Promise((resolve) => {
    const url = `${window.location.origin}${window.location.pathname.replace(/\/[^/]+$/, "/")}dialogs/weekday-range.html?v=${Date.now()}`;

    Office.context.ui.displayDialogAsync(
      url,
      {
        height: 66,
        width: 42,
        displayInIframe: true,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          resolve({
            type: "error",
            message: result.error?.message || "Could not open the weekday range dialog.",
          });
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args: { message: string }) => {
          try {
            const payload = JSON.parse(args.message) as {
              type?: string;
              params?: WeekdayRangeParams;
            };

            if (payload.type === "close") {
              dialog.close();
              resolve({ type: "info", message: "Weekday range dialog closed." });
              return;
            }

            if (payload.type !== "run" || !payload.params) {
              return;
            }

            const computed = buildWeekdayRangeResult(payload.params);
            dialog.messageChild(
              JSON.stringify({
                type: "result",
                result: computed,
              }),
            );
          } catch (error) {
            console.error(error);
            dialog.messageChild(
              JSON.stringify({
                type: "result",
                result: {
                  type: "error",
                  message: "Could not process the weekday range request.",
                  output: "",
                },
              }),
            );
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          resolve({ type: "info", message: "Weekday range dialog closed." });
        });
      },
    );
  });
}

export function openSelectedDeckDialog(): Promise<ActionResult> {
  return new Promise((resolve) => {
    const url = `${window.location.origin}${window.location.pathname.replace(/\/[^/]+$/, "/")}dialogs/selected-deck.html?v=${Date.now()}`;
    let settled = false;

    Office.context.ui.displayDialogAsync(
      url,
      {
        height: 52,
        width: 34,
        displayInIframe: true,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          settled = true;
          resolve({
            type: "error",
            message: result.error.message || "Could not open the Save Selected Deck dialog.",
          });
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args: { message: string }) => {
          let payload: { type?: string } | null = null;
          try {
            payload = JSON.parse(args.message) as { type?: string } | null;
          } catch {
            payload = null;
          }

          if (!payload?.type) {
            return;
          }

          if (payload.type === "close") {
            dialog.close();
            if (!settled) {
              settled = true;
              resolve({ type: "info", message: "Save Selected Deck cancelled." });
            }
            return;
          }

          if (payload.type !== "run") {
            return;
          }

          try {
            const presentationTools = await import("./presentationTools");
            const actionResult = await presentationTools.createDeckFromSelectedSlides();

            if (typeof dialog.messageChild === "function") {
              dialog.messageChild(JSON.stringify({ type: "result", result: actionResult }));
            }
          } catch (error) {
            const failureResult: ActionResult = {
              type: "error",
              message: error instanceof Error ? error.message : "Something went wrong.",
            };

            if (typeof dialog.messageChild === "function") {
              dialog.messageChild(JSON.stringify({ type: "result", result: failureResult }));
            }
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (!settled) {
            settled = true;
            resolve({ type: "info", message: "Save Selected Deck closed." });
          }
        });

        resolve({
          type: "info",
          message: "Opened the Save Selected Deck dialog.",
        });
      },
    );
  });
}

export async function openJolifyWebsite(): Promise<ActionResult> {
  try {
    Office.context.ui.openBrowserWindow("https://stephenjoly.github.io/jolify-ppt/");
    return {
      type: "success",
      message: "Opened the Jolify website.",
    };
  } catch {
    return {
      type: "error",
      message: "Could not open the Jolify website.",
    };
  }
}
// ─────────────────────────────────────────────────────────────────
// Slide organisation
// ─────────────────────────────────────────────────────────────────

const UNUSED_DIVIDER_TAG = "JOLIFY_UNUSED_DIVIDER";
const UNUSED_SECTION_NAME = "Unused Slides";

export async function moveToUnusedSection(): Promise<ActionResult> {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.8")) {
    return {
      type: "error",
      message: "Move to Unused requires PowerPointApi 1.8 (slide move support).",
    };
  }

  return PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    const allSlides = context.presentation.slides;
    selectedSlides.load("items/id");
    allSlides.load("items/id");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return { type: "warning", message: "Select at least one slide in the slide panel first." };
    }

    // Find the tagged divider slide that marks the start of the Unused Slides area.
    const tagChecks = allSlides.items.map((slide) => slide.tags.getItemOrNullObject(UNUSED_DIVIDER_TAG));
    await context.sync();

    const taggedDividerIndexes: number[] = [];
    tagChecks.forEach((tag, index) => {
      if (!tag.isNullObject) {
        taggedDividerIndexes.push(index);
      }
    });

    let dividerSlide: PowerPoint.Slide | null =
      taggedDividerIndexes.length > 0 ? allSlides.items[taggedDividerIndexes[0]] : null;

    // If multiple divider tags exist, keep the first one and remove the rest.
    if (taggedDividerIndexes.length > 1) {
      for (let i = 1; i < taggedDividerIndexes.length; i += 1) {
        allSlides.items[taggedDividerIndexes[i]].tags.delete(UNUSED_DIVIDER_TAG);
      }
      await context.sync();
    }

    // Create the divider slide if it doesn't exist yet.
    let createdSection = false;
    if (!dividerSlide) {
      allSlides.add();
      await context.sync();
      allSlides.load("items/id");
      await context.sync();

      const divSlide = allSlides.items[allSlides.items.length - 1];
      divSlide.tags.add(UNUSED_DIVIDER_TAG, "true");

      const tb = divSlide.shapes.addTextBox(UNUSED_SECTION_NAME, {
        left: 40,
        top: SLIDE_HEIGHT / 2 - 40,
        width: SLIDE_WIDTH - 80,
        height: 60,
      });
      tb.textFrame.textRange.font.size  = 24;
      tb.textFrame.textRange.font.color = "#A0A0A0";
      tb.textFrame.horizontalAlignment  = PowerPoint.ParagraphHorizontalAlignment.center;
      await context.sync();

      createdSection = true;
      dividerSlide = divSlide;
    }

    const dividerId = dividerSlide.id;

    // Keep the divider tag current in case it was manually edited.
    dividerSlide.tags.add(UNUSED_DIVIDER_TAG, "true");
    await context.sync();

    // Build an index map for current positions.
    allSlides.load("items/id");
    await context.sync();
    const idToPos = new Map<string, number>();
    allSlides.items.forEach((slide, index) => idToPos.set(slide.id, index));

    const dividerIdx = idToPos.get(dividerId);
    if (dividerIdx === undefined) {
      return {
        type: "error",
        message: `Could not locate "${UNUSED_SECTION_NAME}" divider slide.`,
      };
    }

    const selectedSlideIds = selectedSlides.items.map((slide) => slide.id).filter((id) => id !== dividerId);
    if (selectedSlideIds.length === 0) {
      return {
        type: "warning",
        message: `Select one or more slides (not the "${UNUSED_SECTION_NAME}" divider slide).`,
      };
    }

    // Only move selected slides that are before the divider (active deck area).
    const slidesToMove = selectedSlideIds
      .filter((id) => {
        const pos = idToPos.get(id);
        return pos !== undefined && pos < dividerIdx;
      })
      .sort((a, b) => (idToPos.get(a) ?? 0) - (idToPos.get(b) ?? 0));

    if (slidesToMove.length === 0) {
      return {
        type: "info",
        message: `Selected slides are already in "${UNUSED_SECTION_NAME}".`,
      };
    }

    // Move each selected slide to just after the divider, preserving relative order.
    let movedCount = 0;
    for (const slideId of slidesToMove) {
      allSlides.load("items/id");
      await context.sync();

      const livePos = new Map<string, number>();
      allSlides.items.forEach((slide, index) => livePos.set(slide.id, index));

      const liveDividerIdx = livePos.get(dividerId);
      const liveSlideIdx = livePos.get(slideId);
      if (liveDividerIdx === undefined || liveSlideIdx === undefined || liveSlideIdx > liveDividerIdx) {
        continue;
      }

      const insertionIdx = Math.min(allSlides.items.length - 1, liveDividerIdx + 1 + movedCount);
      allSlides.getItem(slideId).moveTo(insertionIdx);
      await context.sync();
      movedCount += 1;
    }

    if (movedCount === 0) {
      return {
        type: "info",
        message: `Selected slides are already in "${UNUSED_SECTION_NAME}".`,
      };
    }

    const n = movedCount;
    const createdMsg = createdSection ? ` Created "${UNUSED_SECTION_NAME}" divider slide.` : "";
    return {
      type: "success",
      message: `Moved ${n} slide${n !== 1 ? "s" : ""} to "${UNUSED_SECTION_NAME}".${createdMsg}`,
    };
  });
}
