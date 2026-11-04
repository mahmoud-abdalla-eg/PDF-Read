# Spec Delta

## Purpose

Provides high-performance, document-grade PDF rendering, tabbed navigation, continuous reading view, and search in a lightweight Windows desktop environment.

## ADDED Requirements

### Requirement: Document Rendering and View Controls
The system SHALL render standard PDF documents with continuous scrolling, page zoom controls, and responsive page layout.

#### Scenario: Opening a PDF document
- **WHEN** the user opens a valid PDF file
- **THEN** the system renders the pages with sharp typography and preserves embedded vector graphics and images

#### Scenario: Adjusting document zoom
- **WHEN** the user changes the zoom level via toolbar presets or mouse wheel with Ctrl
- **THEN** the system recalculates page dimensions and re-renders pages without blurriness

### Requirement: Tabbed Document Navigation
The system SHALL support opening multiple PDF documents simultaneously in separate tabs within a single application window.

#### Scenario: Opening multiple files in tabs
- **WHEN** the user opens a second PDF document while an existing document is open
- **THEN** the system creates a new tab displaying the second document without closing the previous tab

#### Scenario: Closing a tab
- **WHEN** the user closes an active tab
- **THEN** the system closes the document view and switches focus to the adjacent active tab

### Requirement: Document Search
The system SHALL provide an in-page text search interface that matches and highlights search terms across the active PDF.

#### Scenario: Finding text occurrences
- **WHEN** the user enters a search query in the search bar
- **THEN** the system highlights matching text and provides next and previous navigation controls

### Requirement: Navigation Sidebar
The system SHALL display a collapsible sidebar containing page thumbnails and document bookmark outlines.

#### Scenario: Navigating by thumbnail or bookmark
- **WHEN** the user clicks a page thumbnail or bookmark outline item
- **THEN** the document view scrolls directly to the corresponding page and target coordinate
