# Spec Delta

## Purpose

Enables two-tier PDF editing comprising in-place non-shifting text block editing with resize controls, markup annotations, and structural page manipulations.

## ADDED Requirements

### Requirement: In-Place Text Block Editing
The system SHALL allow users to select any recognized text block on a PDF page and edit its text in an overlay bounding box without altering adjacent page layout.

#### Scenario: Entering text edit mode
- **WHEN** the user selects the Text Edit tool and clicks on a text line or paragraph
- **THEN** the system overlays an editable bounding box aligned exactly to the text coordinates with font size matching the original text

#### Scenario: Modifying text content
- **WHEN** the user types new text into the active bounding box
- **THEN** the text updates within the bounding box while the surrounding elements on the page remain in their fixed positions

#### Scenario: Resizing bounding box
- **WHEN** the user drags a resize handle on the active text bounding box
- **THEN** the box dimensions update smoothly and re-wrap the contained text within the new boundary

### Requirement: PDF Text Stream Persistence
The system SHALL persist modified text blocks into the underlying PDF stream by redacting the original text region and rendering the new text at the configured coordinates.

#### Scenario: Saving document with in-place text edits
- **WHEN** the user saves the document after modifying text blocks
- **THEN** the system outputs a valid PDF containing the updated text with matching font attributes and the original text cleanly covered

### Requirement: Document Annotations and Markup
The system SHALL support creating, editing, and deleting document annotations including text highlights, notes, freehand drawings, and stamps.

#### Scenario: Adding text highlights
- **WHEN** the user selects the highlight tool and drags across a range of text
- **THEN** the system applies a semi-transparent color highlight over the selected text coordinates

#### Scenario: Adding a sticky note annotation
- **WHEN** the user places a note icon on the page and enters comment text
- **THEN** the system saves the note popup linked to that page coordinate

### Requirement: Page Operations
The system SHALL allow reorganizing document pages including rotation, deletion, reordering, and inserting pages from other PDF files.

#### Scenario: Rotating a page
- **WHEN** the user chooses to rotate the current page clockwise by 90 degrees
- **THEN** the system updates the page rotation dictionary and refreshes the page display

#### Scenario: Deleting or reordering pages
- **WHEN** the user deletes a selected page or drags a thumbnail to a new index
- **THEN** the system updates the page tree structure and reflects the new sequence
