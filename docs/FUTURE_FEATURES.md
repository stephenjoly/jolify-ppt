# Future Features

This file tracks Jolify features that are intentionally out of the current ribbon build and still need to be designed and implemented properly.

1. Remove Presentation Comments
Description:
Remove all slide comments from the current presentation and produce a cleaned copy that is safe to share.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Jolify creates a new deck without slide comments.
- Original presentation is left unchanged.
- User gets a clear success or no-op message.
- If no comments exist, Jolify reports that nothing was removed.

2. Copy Presentation Comments
Description:
Extract all presentation comments and copy them to the clipboard in a readable review format.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Copied output includes slide-level grouping.
- Clipboard output is formatted consistently for review use.
- If no comments exist, Jolify reports that nothing was copied.

3. Download Presentation Comments
Description:
Export all presentation comments to a Markdown file for offline review or archival.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Exported file is Markdown.
- File content is grouped clearly by slide.
- Filename is sensible and presentation-based.
- If no comments exist, Jolify reports that nothing was exported.

4. Remove Speaker Notes
Description:
Remove all speaker notes from the current presentation and generate a clean copy for external sharing.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Jolify creates a new deck without speaker notes.
- Original presentation is left unchanged.
- User gets a clear success or no-op message.
- If no notes exist, Jolify reports that nothing was removed.

5. Copy Speaker Notes
Description:
Extract all speaker notes and copy them to the clipboard in a readable structured format.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Copied output is organized by slide.
- Clipboard output is review-friendly and consistent.
- If no notes exist, Jolify reports that nothing was copied.

6. Download Speaker Notes
Description:
Export all speaker notes to a Markdown file for editing, review, or archival.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Exported file is Markdown.
- File content is grouped clearly by slide.
- Filename is sensible and presentation-based.
- If no notes exist, Jolify reports that nothing was exported.

7. Save Selected Slides As New Deck
Description:
Create a new presentation from the currently selected slides and save or download it as its own deck.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Action requires one or more selected slides.
- Output deck contains only the selected slides.
- Saved filename is sensible and presentation-based.
- User gets a clear error if no slides are selected.

8. Attach Selected Slides To Email
Description:
Create a new deck from selected slides and attach it to a new email draft.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Action requires one or more selected slides.
- Generated deck is attached to a new draft email.
- Local-mode dependency is handled explicitly when required.
- User gets a clear error if no slides are selected.

9. Client-Ready Export
Description:
Create a copy of the current deck with all hidden slides removed so the presentation is ready for client sharing.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Jolify creates a new deck with hidden slides removed.
- Original presentation is left unchanged.
- User gets a clear count of removed slides.
- If no hidden slides exist, Jolify reports that nothing was removed.

10. Gantt Chart Builder
Description:
Provide a supported Gantt creation workflow that can generate and update timeline-style Gantt visuals in PowerPoint.
Key acceptance criteria:
- User can launch the workflow from the ribbon.
- Builder supports creating a new Gantt from structured inputs.
- Builder supports editing an existing Jolify-created Gantt.
- Output is stable in normal slide-editing workflows.
- User input and update behavior are documented clearly enough to support maintenance.

## Notes

- These features are intentionally out of the current stable ribbon.
- Ribbon reintegration should happen only after each feature is verified end-to-end.
- Native save and email workflows may require local mode on macOS.
- Hosted mode should prefer download-based fallbacks when native flows are unavailable.
