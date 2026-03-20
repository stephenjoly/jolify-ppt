# Unsupported Native Control Notes

These PowerPoint ribbon integrations were tested in Jolify on macOS and should be treated as unreliable or unsupported for this project unless Microsoft changes host behavior.

## Tested Environment

- PowerPoint for Mac `16.107.1`
- `AddinCommands 1.3` declared in the manifest
- full `npm stop` / `npm start` cycle
- full PowerPoint restart after manifest changes

## Do Not Revisit By Default

### Whole native Office groups

These were added as `OfficeGroup` elements and did not render on the Jolify tab:

- `GroupFont`
- `Paragraph`
- `GroupSlides`

Conclusion:
- individual built-in `OfficeControl` items can work on this host
- whole built-in `OfficeGroup` injection is not reliable for Jolify on this Mac build

### Native paragraph dialog control

This was tested as a standalone built-in control and still did not surface reliably:

- `PowerPointParagraphDialog`

Conclusion:
- do not plan around a dedicated paragraph-dialog button in the ribbon
- if paragraph spacing must be edited, use PowerPoint's native Home workflow instead

### Standalone native reset-slide button

We did not find a verified supported standalone built-in control for Reset Slide, and the broader slide group injection did not solve it.

Conclusion:
- keep `SlideLayoutGallery` as the working layout control
- do not spend more time trying to expose Reset Slide through guessed control IDs

### Standalone native regroup button

We did not find a supported standalone built-in control ID for Regroup in the Office control surface used by the add-in.

Conclusion:
- do not spend more time trying random regroup IDs in the manifest

## API Limits

### Paragraph spacing before/after

PowerPoint Office.js currently exposes `TextRange.paragraphFormat`, but the usable paragraph properties available here are alignment and bullets, not paragraph spacing before/after.

Conclusion:
- do not plan a Jolify button that directly sets paragraph spacing before/after through Office.js

## Working Pattern

Prefer these instead:

- individual built-in `OfficeControl` items that have already rendered successfully
- custom Jolify commands backed by `src/shared/shapeTools.ts`
- native PowerPoint workflows for paragraph spacing and other text features the API does not expose
