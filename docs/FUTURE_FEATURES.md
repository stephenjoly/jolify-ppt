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

11. Stretch Shapes To Reference Edge
Description:
Provide edge-matching actions where the first selected shape is the reference and the other selected shape(s) are extended or shrunk so their ending edge lands on the same left, right, top, or bottom position as the reference shape.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- Action requires a reference shape plus at least one additional target shape.
- First selected shape is treated as the reference edge source.
- Target shape(s) keep their opposite edge fixed and resize so the chosen edge matches the reference shape.
- Jolify supports the expected directional variants clearly, such as stretch to match right edge and equivalent left/top/bottom versions.
- User gets a clear error if the selection is missing, invalid, or cannot be resized as requested.

The ribbon convenience items below are ordered from likely simplest to more involved implementation.

12. Arrange And Group Shortcuts
Description:
Add direct ribbon buttons for common native shape-arrangement actions, including bring to front, send to back, group, ungroup, and regroup, so users do not need to leave the Jolify tab for routine stacking and grouping work.
Key acceptance criteria:
- User can trigger each action directly from the Jolify ribbon.
- Buttons map cleanly to the expected native PowerPoint behavior.
- Actions work on valid selections without requiring the taskpane.
- Jolify gives a clear error or disabled state when the current selection cannot use the requested action.

13. Layout And Reset Shortcuts
Description:
Add direct ribbon buttons for slide layout selection and reset slide so common cleanup and template-restoration actions are available from the Jolify tab.
Key acceptance criteria:
- User can trigger layout and reset from the Jolify ribbon.
- Actions apply to the current slide selection using normal PowerPoint behavior.
- Jolify does not override or reinterpret native placeholder/layout logic.
- User gets a clear error or disabled state when the action is not applicable.

14. Add Page Numbers Across Deck
Description:
Add a ribbon entrypoint for enabling slide page numbers across the current presentation so numbering can be applied to every slide without switching back to native tabs.
Key acceptance criteria:
- User can trigger the action from the Jolify ribbon.
- The workflow uses expected PowerPoint page-number behavior rather than a custom numbering system.
- Page numbers are applied across the deck where PowerPoint and the active master support them.
- Result is consistent with the active theme or master where PowerPoint supports it.
- User gets a clear success, no-op, or guidance message if numbering cannot be applied as expected.

15. Comment Shortcuts
Description:
Add ribbon buttons for add comment, next comment, and previous comment so users can move through review comments from the Jolify tab.
Key acceptance criteria:
- User can trigger add comment, next comment, and previous comment directly from the Jolify ribbon.
- Commands follow native PowerPoint comment behavior and navigation order.
- Buttons behave clearly when a deck has no comments or comment navigation is unavailable.
- Jolify does not interfere with existing review-mode behavior.

16. Add Shape And Edit Shape Shortcuts
Description:
Add ribbon buttons for add shape and edit shape so common shape-creation and native shape-edit workflows are accessible from the Jolify tab.
Key acceptance criteria:
- User can launch add shape and edit shape from the Jolify ribbon.
- Actions hand off to native PowerPoint shape workflows rather than creating a separate Jolify drawing model.
- Edit shape is only enabled when the current selection supports it.
- User gets a clear error or disabled state when no compatible shape is selected.

17. Draw Shortcut
Description:
Add a ribbon button for the PowerPoint draw command so users can enter the drawing workflow from the Jolify tab when markup or freehand annotation is needed.
Key acceptance criteria:
- User can trigger the draw command from the Jolify ribbon.
- Action opens or switches into the expected native drawing workflow.
- Jolify makes it clear whether the button is a launcher, mode switch, or tab shortcut.
- Behavior is verified in the supported PowerPoint environments because draw features can vary by platform and license state.

18. Normalize Slide Title Case
Description:
Add a bulk text action that updates the title on every slide to a chosen casing style, such as title case, uppercase, or sentence case, in one run.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- User can choose the target casing style before the update runs.
- Jolify updates slide-title text across the presentation in one pass.
- The action targets title placeholders or clearly defined title shapes only, rather than rewriting arbitrary body text.
- User gets a clear report of how many slide titles were updated and which slides were skipped.

19. Normalize Subtitle Ending Periods
Description:
Add a bulk text action that enforces subtitle punctuation consistently across the presentation by either adding a trailing period to every subtitle or removing trailing periods from every subtitle.
Key acceptance criteria:
- User can trigger the action from the ribbon.
- User can choose between always adding a trailing period or always removing it.
- Jolify updates subtitle text across the presentation in one pass.
- The action targets subtitle placeholders or clearly defined subtitle shapes only, rather than rewriting arbitrary text boxes.
- User gets a clear report of how many subtitles were updated and which slides were skipped.

## Notes

- These features are intentionally out of the current stable ribbon.
- Ribbon reintegration should happen only after each feature is verified end-to-end.
- Native save and email workflows may require local mode on macOS.
- Hosted mode should prefer download-based fallbacks when native flows are unavailable.
