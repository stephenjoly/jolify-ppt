# Jolify vs. Templafy Taskpane

This document compares Jolify's current taskpane and ribbon surface against the Templafy Productivity tools pane feature list the user provided on March 22, 2026.

Scope notes:
- This is a product-surface comparison, not an implementation-difficulty estimate.
- `Jolify` here means the current live add-in surface across the ribbon and taskpane, not just what is visible in one screenshot.
- Status labels:
  - `Available`: Jolify has a reasonably direct equivalent today.
  - `Partial`: Jolify covers some of the workflow, but not the full Templafy feature set in that section.
  - `Missing`: no current equivalent in Jolify.
  - `Different`: Jolify solves the problem in a meaningfully different way rather than matching the exact Templafy control.

## Overall Read

Jolify is currently strongest in:
- geometry copying and matching
- alignment and distribution
- size matching and edge stretching
- text-box cleanup
- a few custom deck utilities such as `Move to Unused`, `Save Selected Deck`, `Toggle Draft`, `Weekday Range`, and `Create Grid`

Templafy is currently much broader in:
- insert primitives and symbol libraries
- line and outline editing
- rotate / flip
- comments / stickers workflows
- picture tooling
- export / share workflows

The biggest philosophical difference is this:
- Templafy is a broad presentation utility shelf.
- Jolify is currently a narrower but more custom geometry and cleanup toolkit.

## Section Summary

| Templafy Section | Jolify Status | Notes |
| --- | --- | --- |
| Insert | Partial | Jolify has `Shapes Insert Gallery`, `Create Grid`, and `Create Center Sticker`, but not Templafy's broader insert catalog. |
| Outline | Partial | Jolify can match/copy/paste outline styling and clear outlines, but lacks dedicated line editing controls like weight, dash style, arrow style, and line normalization. |
| Align | Partial | Jolify covers core align/distribute strongly, but not select-similar, distribute-with-resize, stack, unify adjustments, or super-size. |
| Size | Available | Jolify has same width, same height, same size, and all four stretch-to-edge actions. |
| Arrange | Partial | Jolify has `Bring to Front` and `Send to Back`, but not group/ungroup or forward/backward stepping. |
| Rotate | Missing | No current rotate / flip tools. |
| Swap | Partial | Jolify has full position swap, but not X-only or Y-only swap. |
| Text | Partial | Jolify has a solid text-box cleanup subset, but not Templafy's full text utility coverage. |
| Picture | Missing | No current picture workflow tooling. |
| Stickers | Partial | Jolify has `Toggle Draft Sticker`, but not Templafy's comments/stickers navigation and cleanup set. |
| Other | Partial | Jolify covers parts of the shape-format pick/apply workflow plus `Move to Unused`, but not workspace/master/powershape equivalents. |
| Share | Partial | Jolify can save selected slides as a new deck, but does not yet match Templafy's export/email matrix. |

## Detailed Comparison

### Insert

Templafy includes:
- insert text box
- insert rectangle / rounded rectangle
- insert flow / arrow / oval / line / elbow line
- insert symbol
- insert Harvey ball
- insert placeholder

Jolify today:
- `Shapes Insert Gallery` on the ribbon
- `Create Grid`
- `Create Center Sticker`
- `Merge Text Boxes`
- `Split Text Box`

Assessment:
- `Partial`

Notes:
- Jolify can reach standard shapes through the native PowerPoint shape gallery, which covers part of the insert use case.
- Jolify does not yet have specialized insert helpers like Harvey balls, placeholders, symbol insertion, or flow drawing helpers.

### Outline

Templafy includes:
- outline weight
- dash style
- arrow style
- make lines vertical
- make lines horizontal

Jolify today:
- `Match Outline`
- `Copy Outline`
- `Paste Outline`
- `No Outline`
- native `Shape Outline Color Picker`

Assessment:
- `Partial`

Notes:
- Jolify is stronger at reusing outline formatting between shapes than at directly editing line properties.
- Templafy's direct line-edit controls do not currently exist in Jolify.

### Align

Templafy includes:
- align left / right / top / center / middle / bottom
- select similar
- distribute horizontally / vertically
- distribute with resize
- stack horizontally / vertically
- unify adjustments
- super size

Jolify today:
- `Align Left`
- `Align Center`
- `Align Right`
- `Align Top`
- `Align Middle`
- `Align Bottom`
- `Distribute H`
- `Distribute V`
- `Center + Group`
- `Middle + Group`
- `Center + Middle + Group`
- `Distribute H + Group`
- `Distribute V + Group`
- taskpane `Distribute H + V + Group`

Assessment:
- `Partial`

Notes:
- Jolify covers the core alignment layer well.
- Jolify does not currently cover the more advanced Templafy align variants like resize-distribute, stack, select-similar, or unify-adjustments.

### Size

Templafy includes:
- same width
- same height
- same size
- stretch left / right / top / bottom

Jolify today:
- `Match Width`
- `Match Height`
- `Match H+W`
- `Match Left Edge`
- `Match Right Edge`
- `Match Top Edge`
- `Match Bottom Edge`

Assessment:
- `Available`

Notes:
- This is one of Jolify's best parity areas.
- Jolify's edge-stretch workflow is a strong fit for real slide work and already feels product-distinctive.

### Arrange

Templafy includes:
- group / ungroup
- send to back / bring to front
- send backward / bring forward

Jolify today:
- `Bring to Front`
- `Send to Back`

Assessment:
- `Partial`

Notes:
- Jolify only covers the two endpoints today.
- Mid-stack ordering and grouping controls are still missing.

### Rotate

Templafy includes:
- rotate right 90
- rotate left 90
- flip vertically
- flip horizontally
- more rotation options

Jolify today:
- no equivalent

Assessment:
- `Missing`

### Swap

Templafy includes:
- swap position
- swap X-position
- swap Y-position

Jolify today:
- `Swap Positions`

Assessment:
- `Partial`

Notes:
- Jolify has the most useful full swap, but not the axis-specific variants.

### Text

Templafy includes:
- default textbox format
- text margins
- auto-size
- word-wrap
- increase / decrease line spacing
- split text box
- merge text boxes
- split table to text boxes
- swap text
- footnotes
- set language

Jolify today:
- `Merge Text Boxes`
- `Split Text Box`
- `Remove Text Margins`
- `AutoFit Off`
- native font / size / bullets / numbering / line spacing / highlight on the ribbon

Assessment:
- `Partial`

Notes:
- Jolify is already useful for text-box cleanup.
- It does not yet match Templafy's broader text utilities, especially table splitting, text swapping, footnotes, or language tools.

### Picture

Templafy includes:
- compress pictures
- crop fill fit
- create presentation from pictures

Jolify today:
- no equivalent

Assessment:
- `Missing`

### Stickers

Templafy includes:
- add comments
- next comment
- remove all comments
- add stickers
- next stickers
- remove all stickers

Jolify today:
- `Toggle Draft Sticker`

Assessment:
- `Partial`

Notes:
- Jolify has one deck-wide sticker/branding action.
- It does not yet have Templafy-style comments or sticker management workflows.

### Other

Templafy includes:
- pick position
- pick size
- pick size and position
- pick other formatting
- apply formatting
- workspace
- slide master view
- send to unused
- make PowerShape

Jolify today:
- `Copy Position`
- `Paste Position`
- `Copy Size`
- `Paste Size`
- `Copy Position & Size`
- `Paste Position & Size`
- `Copy Fill / Paste Fill`
- `Copy Outline / Paste Outline`
- `Copy Style / Paste Style`
- `Match Fill`
- `Match Outline`
- `Match Style`
- `Move to Unused`

Assessment:
- `Partial`

Notes:
- Jolify compares well on pick/apply geometry and formatting.
- It does not currently cover workspace, slide master view, or PowerShape-style object construction.

### Share

Templafy includes:
- clean-up
- export presentation as PPTX / PDF
- export selected slides as PPTX / PDF
- email presentation / slides as PPTX / PDF

Jolify today:
- `Save Selected Deck`

Assessment:
- `Partial`

Notes:
- Jolify can already produce a new deck from selected slides, which overlaps part of the selected-slides export story.
- PDF, email, presentation-wide export variants, and cleanup workflows are still missing.

## Current Jolify Advantages

Areas where Jolify already feels especially strong or distinctive:
- `Stretch Shapes To Reference Edge`
- style clipboard plus match-style workflows
- combined align-and-group actions
- `Create Grid`
- `Weekday Range`
- `Move to Unused`
- `Save Selected Deck`
- `Create Center Sticker`

These are not one-to-one Templafy clones; they are Jolify-specific utilities with strong day-to-day slide value.

## Highest Gaps If The Goal Is "Templafy-Like"

If the goal is to close the biggest visible gaps first, the most obvious missing clusters are:

1. `Rotate`
2. `Arrange` depth: group / ungroup / bring forward / send backward
3. `Insert` breadth: symbols, Harvey balls, placeholders, line helpers
4. `Picture` tools
5. `Comments / Stickers` workflows
6. `Share / Export` matrix, especially PDF and email flows

## Taskpane UI Parity Notes

Compared with Templafy's pane, Jolify now has:
- compact sectioned icon layout
- hover help text
- category grouping

But it still differs in a few visible ways:
- some Jolify taskpane actions use monogram placeholders instead of polished icons
- Templafy's visual language is flatter and more uniform
- Jolify has fewer sections overall and a more custom-utility-heavy mix

So the current state is:
- functionality parity in geometry is decent
- functionality parity across the whole Templafy pane is still far away
- UI parity is improving, but not there yet
