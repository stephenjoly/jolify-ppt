# Templafy Gap Status

This file records the remaining gaps from `docs/TEMPLAFY_SAMPLE_ICONS_REFERENCE.md` and separates them into:

- hard-blocked by the current Office.js or host surface
- feasible later, but larger than the current safe incremental work
- partially built already, with a clear next refinement if we choose to continue

The goal is to stop revisiting API-dead ends without evidence that the platform changed.

## Hard-Blocked Or Host-Limited

These are not good candidates for immediate implementation in the current Jolify model.

### PowerPoint Office.js API blockers

- `Increase line spacing`
- `Decrease line spacing`
  - PowerPoint `ParagraphFormat` in Office.js exposes alignment and bullets, not line spacing values.

- `Outline arrow style`
  - PowerPoint `ShapeLineFormat` does not expose arrowhead properties, even though Excel does.

- `Flip vertically`
- `Flip horizontally`
  - PowerPoint Office.js exposes `rotation`, but not a true flip/mirror operation.

- `Slide Master View`
  - There is no supported Office.js surface to toggle PowerPoint into Slide Master View.

- `Set language`
  - There is no verified PowerPoint Office.js surface for opening or setting proofing language.

### Native PowerPoint surface blockers already tested

- native font / paragraph / comments / drawing Office groups in the custom ribbon
- native paragraph dialog control
- native reset-slide button
- native regroup button

These were already tested and documented in [UNSUPPORTED_NATIVE_CONTROLS.md](/Users/stephenjoly/Documents/Coding/ppt-addin/docs/UNSUPPORTED_NATIVE_CONTROLS.md).

### Currently flaky local-automation path

- `Selected slides as PDF`
  - Conceptually possible through the local bridge, but the current PowerPoint AppleScript export-to-PDF path was not reliable enough to ship.

## Feasible But Larger Product Work

These are not blocked by a single missing API, but they are bigger workflows and should be treated as real feature projects.

- `Insert flow`
  - Needs a custom multi-step chevron builder, not a simple insert button.

- `Insert placeholder`
  - Needs slide-layout and placeholder management logic, likely across multiple selected slides.

- `Select similar`
  - Needs a shape-comparison engine with explicit matching rules.

- `Unify adjustments`
  - Needs shape-specific geometry adjustment handling beyond normal width/height matching.

- `Super size`
  - Needs coordinated scaling of shape size, text size, margins, and outlines.

- `Split table to text boxes`
  - Needs table extraction and shape generation logic that preserves row/column structure reasonably well.

- `Footnotes`
  - Needs a full text-reference and footnote management model, not just a button.

- `Compress pictures`
  - Likely requires local-mode image processing and PPTX package updates, or native PowerPoint automation.

- `Crop / Fill / Fit`
  - Needs picture-cropping behavior that PowerPoint Office.js does not expose directly.

- `Work Space`
  - Needs a defined Jolify concept of workspace overlay and slide-wide state management.

- `Make PowerShape`
  - Needs a full reusable-shape definition workflow.

- `Clean-up` (partially built: comments, notes, hidden slides, unused slides, and section stripping)
  - Feasible, but should be designed as a bounded redaction/export workflow rather than a single command.

- `Email presentation as PPTX`
- `Email presentation as PDF`
- `Email selected slides as PPTX`
- `Email selected slides as PDF`
  - Feasible through the local bridge, but each one is a real workflow with attachment generation, local save handling, and Outlook draft creation.

## Partially Built

These now exist, but are still narrower than the Templafy version.

- `Insert Harvey ball`
  - Current state: inserts a bounded quarter-filled Harvey ball.
  - Missing for fuller parity: fraction variants and state cycling.

- `Create presentation from pictures`
  - Current state: local-mode multi-image import, one fitted image per slide.
  - Missing for fuller parity: true folder-based workflow and any broader image-format handling beyond the current supported set.

## Still Feasible Later

These are not blocked in principle, but they were lower priority than the recent clean wins.

- `Email presentation as PPTX`
  - Strongest next email candidate because full-deck PPTX export already exists.

- `Clean-up`
  - Strongest next bounded workflow candidate if we want a meaningful “share-prep” feature.

- `Insert flow`
  - Good custom-Jolify candidate if we want to keep building slide-construction tools.

- `Split table to text boxes`
  - Probably the most plausible next non-export advanced editing feature.

## Recommendation

If we continue closing gaps, the best remaining non-blocked items are:

1. `Clean-up`
   Current Jolify scope removes comments, speaker notes, hidden slides, the Jolify unused-slides area, and section metadata in a new cleaned copy. Still deferred: metadata stripping, dummy-text replacement/redaction, and destructive in-place cleanup.
2. `Email presentation as PPTX`
3. `Insert flow`
4. `Split table to text boxes`
5. richer `Harvey ball` behavior

If we want to avoid bigger projects, then we are effectively done with the clean low-risk parity set and should stop forcing the blocked items.
