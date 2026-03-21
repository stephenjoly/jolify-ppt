#!/usr/bin/env node
/**
 * Generates ribbon button icons from Fluent UI System Icons.
 * Outputs PNG files at 16, 32, and 80px into assets/icons/.
 *
 * Run: node scripts/generate-icons.js
 */

const fs = require("fs");
const path = require("path");
const { Resvg } = require("@resvg/resvg-js");

const ICONS_SRC = path.join(__dirname, "../node_modules/@fluentui/svg-icons/icons");
const OUT_DIR = path.join(__dirname, "../assets/icons");
const SIZES = [16, 32, 80];

// Maps each ribbon button ID → Fluent UI icon filename (without .svg)
const ICON_MAP = {
  // Position
  copyPositionAndSize:  "copy_20_regular",
  pastePositionAndSize: "clipboard_paste_20_regular",
  copyPositionOnly:     "my_location_20_regular",
  pastePositionOnly:    "location_arrow_20_regular",
  copySizeOnly:         "resize_20_regular",
  pasteSizeOnly:        "arrow_autofit_content_20_regular",
  swapPositions:        "arrow_swap_20_regular",
  // Match
  copyOutlineStyle:     "border_outside_20_regular",
  copyFillStyle:        "color_background_20_regular",
  clearFill:            "paint_brush_subtract_20_regular",
  clearOutline:         "border_none_20_regular",
  matchHeight:          "arrow_autofit_height_20_regular",
  matchWidth:           "arrow_autofit_width_20_regular",
  matchHeightAndWidth:  "arrow_autofit_content_20_regular",
  // Align
  alignLeft:            "align_left_20_regular",
  alignCenterH:         "align_center_horizontal_20_regular",
  alignRight:           "align_right_20_regular",
  alignTop:             "align_top_20_regular",
  alignMiddleV:         "align_center_vertical_20_regular",
  alignBottom:          "align_bottom_20_regular",
  distributeH:          "text_align_distributed_evenly_20_regular",
  distributeV:          "text_align_distributed_vertical_20_regular",
  distributeHandVAndGroup: "text_align_distributed_20_regular",
  openWeekdayRangeDialog: "calendar_ltr_20_regular",
  openJolifyWebsite:    "globe_desktop_20_regular",
  // Text
  splitTextBoxByLines:  "text_column_two_20_regular",
  removeTextMargins:    "textbox_align_top_left_20_regular",
  disableTextAutofit:   "text_wrap_off_20_filled",
  createCenterSticker:  "sticker_20_regular",
  // Branding
  addDraftSticker:      "tag_add_20_regular",
  removeDraftSticker:   "tag_dismiss_20_regular",
};

fs.mkdirSync(OUT_DIR, { recursive: true });

let ok = 0, warn = 0;

for (const [button, iconName] of Object.entries(ICON_MAP)) {
  const svgPath = path.join(ICONS_SRC, `${iconName}.svg`);
  if (!fs.existsSync(svgPath)) {
    console.warn(`  ⚠  Missing source icon: ${iconName}  (button: ${button})`);
    warn++;
    continue;
  }

  const svgData = fs.readFileSync(svgPath, "utf-8");

  for (const size of SIZES) {
    const resvg = new Resvg(svgData, {
      fitTo: { mode: "width", value: size },
    });
    const pngBuffer = resvg.render().asPng();
    const outFile = path.join(OUT_DIR, `${button}_${size}.png`);
    fs.writeFileSync(outFile, pngBuffer);
    ok++;
  }

  console.log(`  ✓  ${button}  ←  ${iconName}`);
}

console.log(`\n  Generated ${ok} PNG files in assets/icons/`);
if (warn) console.warn(`  ${warn} icon(s) missing — check ICON_MAP in scripts/generate-icons.js`);
