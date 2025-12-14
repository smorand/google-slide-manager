# Google Slide Manager

Comprehensive command-line tool for managing Google Slides presentations with support for creation, editing, formatting, translation, and export operations.

## Features

### Presentation Operations
- Create new presentations
- Move presentations to specific folders

### Slide Operations
- Add new slides with custom layouts
- Duplicate existing slides
- Remove slides
- Reorder multiple slides
- Move slides to new positions

### Table Operations
- Create tables on slides
- Update table cell content
- Style table cells with background colors

### Text Operations
- Find and replace text across presentations
- Extract all text from presentations
- Search for specific text within presentations

### Notes Operations
- Get speaker notes from slides
- Add speaker notes to slides
- Extract all notes from presentations

### Shape Operations
- Add shapes to slides (rectangles, ellipses, etc.)

### Style Operations
- Copy text styles between elements
- Copy themes between presentations

### Translation
- Translate slides to target languages

### Export
- Export presentations as PDF
- Export presentations as PowerPoint (PPTX)

## Installation

### Prerequisites
- Go 1.21 or higher
- Google Cloud credentials with Slides and Drive API access

### Build and Install

```bash
# Build the binary
make build

# Install to /usr/local/bin
make install

# Or install to a custom location
TARGET=/path/to/bin make install
```

### Uninstall

```bash
make uninstall
```

## Setup

### 1. Google Cloud Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Slides API and Google Drive API
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials JSON file
6. Save it as `~/.credentials/google_credentials.json`

### 2. First Run

On first run, the tool will prompt you to authenticate:
```bash
google-slide-manager create-presentation "My First Presentation"
```

Follow the authorization URL and enter the code when prompted. The token will be saved to `~/.credentials/google_token.json`.

## Usage

### Presentation Operations

#### Create Presentation
```bash
# Create a new presentation
google-slide-manager create-presentation "My Presentation"

# Create in specific folder
google-slide-manager create-presentation "My Presentation" --folder FOLDER_ID
```

### Slide Operations

#### Add Slide
```bash
# Add a blank slide
google-slide-manager add-slide PRESENTATION_ID

# Add slide with specific layout
google-slide-manager add-slide PRESENTATION_ID --layout TITLE

# Available layouts: BLANK, TITLE, TITLE_AND_BODY, TITLE_ONLY, etc.
```

#### Duplicate Slide
```bash
google-slide-manager duplicate-slide PRESENTATION_ID SLIDE_INDEX
```

#### Remove Slide
```bash
google-slide-manager remove-slide PRESENTATION_ID SLIDE_INDEX
```

#### Move Slide
```bash
google-slide-manager move-slide PRESENTATION_ID SLIDE_INDEX NEW_POSITION
```

#### Reorder Slides
```bash
# Reorder slides by providing comma-separated indices
google-slide-manager reorder-slides PRESENTATION_ID "2,0,1,3"
```

### Table Operations

#### Create Table
```bash
google-slide-manager create-table PRESENTATION_ID SLIDE_INDEX ROWS COLS
```

#### Update Cell
```bash
google-slide-manager update-cell PRESENTATION_ID TABLE_ID ROW COL "Cell Text"
```

#### Style Cell
```bash
google-slide-manager style-cell PRESENTATION_ID TABLE_ID ROW COL --bg-color "#FF0000"
```

### Text Operations

#### Replace Text
```bash
google-slide-manager replace-text PRESENTATION_ID "find text" "replacement text"
```

#### Extract All Text
```bash
google-slide-manager extract-all-text PRESENTATION_ID
```

#### Search Text
```bash
google-slide-manager search-text PRESENTATION_ID "search query"
```

### Notes Operations

#### Get Notes
```bash
google-slide-manager get-notes PRESENTATION_ID SLIDE_INDEX
```

#### Add Notes
```bash
google-slide-manager add-notes PRESENTATION_ID SLIDE_INDEX "Speaker notes here"
```

#### Extract All Notes
```bash
google-slide-manager extract-all-notes PRESENTATION_ID
```

### Shape Operations

#### Add Shape
```bash
# Available shapes: RECTANGLE, ELLIPSE, etc.
google-slide-manager add-shape PRESENTATION_ID SLIDE_INDEX RECTANGLE
```

### Export Operations

#### Export as PDF
```bash
google-slide-manager export-pdf PRESENTATION_ID output.pdf
```

#### Export as PowerPoint
```bash
google-slide-manager export-pptx PRESENTATION_ID output.pptx
```

## Development

### Build
```bash
make build
```

### Clean
```bash
# Remove build artifacts
make clean

# Remove all including go.mod and go.sum
make clean-all
```

### Rebuild
```bash
make rebuild
```

### Format Code
```bash
make fmt
```

### Run Tests
```bash
make test
```

### Run Checks
```bash
# Run fmt, vet, and test
make check
```

## Project Structure

```
google-slide-manager/
├── Makefile              # Build automation
├── README.md             # Human-oriented documentation
├── CLAUDE.md             # AI-oriented documentation
└── src/                  # Go source code
    ├── main.go           # Main entry point
    ├── cli.go            # CLI command implementations
    └── auth.go           # Google OAuth authentication
```

## API Scopes

The tool requires the following Google API scopes:
- `https://www.googleapis.com/auth/presentations` - Google Slides API
- `https://www.googleapis.com/auth/drive.file` - Google Drive API (for file operations)
- `https://www.googleapis.com/auth/cloud-translation` - Translation API (for translation features)

## License

This project is provided as-is for managing Google Slides presentations.

## Troubleshooting

### Authentication Issues
If you encounter authentication errors:
1. Delete `~/.credentials/google_token.json`
2. Run any command again to re-authenticate

### API Errors
- Ensure the Google Slides API and Drive API are enabled in your Google Cloud project
- Verify your credentials file is valid and in the correct location

### Permission Errors
When installing:
```bash
# May require sudo for /usr/local/bin
sudo make install

# Or use a directory in your PATH that doesn't require sudo
TARGET=~/bin make install
```
