# Google Slide Manager - AI Documentation

## Project Overview

Command-line tool for comprehensive Google Slides management built in Go using the Google Slides API and Drive API. Provides programmatic access to presentation creation, editing, formatting, and export operations.

## Architecture

### Technology Stack
- **Language**: Go 1.21+
- **CLI Framework**: cobra (github.com/spf13/cobra)
- **Google APIs**:
  - Google Slides API v1 (google.golang.org/api/slides/v1)
  - Google Drive API v3 (google.golang.org/api/drive/v3)
- **Authentication**: OAuth2 (golang.org/x/oauth2)
- **Output Formatting**: fatih/color for terminal colors

### Project Structure

Following **Standard Go Project Layout**:

```
google-slide-manager/
├── Makefile                       # Build automation
├── README.md                      # Human-oriented documentation
├── CLAUDE.md                      # This file - AI-oriented documentation
├── go.mod                         # Go module definition
├── go.sum                         # Dependency checksums
├── cmd/                           # Main applications
│   └── google-slide-manager/      # Binary entry point
│       └── main.go                # Minimal entry point (wiring only)
└── internal/                      # Private application code
    ├── auth/                      # Authentication package
    │   └── auth.go                # OAuth2 client and service creation
    ├── cli/                       # CLI command definitions
    │   └── cli.go                 # Cobra commands and flag management
    ├── presentation/              # Presentation operations
    │   └── presentation.go        # Create and get presentations
    ├── slide/                     # Slide operations
    │   └── slide.go               # Add, duplicate, move, remove, reorder slides
    ├── table/                     # Table operations
    │   └── table.go               # Create tables, update cells, apply styles
    ├── text/                      # Text operations
    │   └── text.go                # Extract, replace, search text
    ├── notes/                     # Speaker notes operations
    │   └── notes.go               # Get, add, extract notes
    ├── shape/                     # Shape operations
    │   └── shape.go               # Add shapes to slides
    ├── style/                     # Style operations
    │   └── style.go               # Copy text styles, themes, translations (placeholders)
    └── export/                    # Export operations
        └── export.go              # Export to PDF and PPTX
```

## Package Descriptions

### cmd/google-slide-manager/main.go
**Purpose**: Minimal application entry point

**Key Components**:
- `main()`: Calls `cli.Execute()` and handles errors

**Design**: Following Go best practices, main.go contains only initialization and error handling. All business logic is in `internal/` packages.

### internal/auth/auth.go
**Purpose**: OAuth2 authentication and service creation

**Exported Functions**:
- `GetCredentialsPath()`: Returns path to credentials directory
- `GetClient(ctx)`: Creates authenticated HTTP client
- `GetSlidesService(ctx)`: Creates Slides API service
- `GetDriveService(ctx)`: Creates Drive API service

**Credentials Location**:
- Credentials: `~/.gdrive/credentials.json`
- Token: `~/.gdrive/token.json`

**API Scopes**:
```go
scopes = []string{
    slides.PresentationsScope,
    drive.DriveScope,
    "https://www.googleapis.com/auth/cloud-translation",
}
```

### internal/cli/cli.go
**Purpose**: CLI command definitions and flag management

**Structure**:
- Package-level flag variables (grouped by command)
- Command initialization functions (initXxxCommands)
- Command definitions (cobra.Command structs)
- Command execution functions (runXxx functions)
- Helper functions (printJSON)

**Command Categories**:
1. **Presentation Operations**: create-presentation
2. **Slide Operations**: add-slide, duplicate-slide, remove-slide, move-slide, reorder-slides
3. **Table Operations**: create-table, update-cell, style-cell
4. **Text Operations**: replace-text, extract-all-text, search-text
5. **Notes Operations**: get-notes, add-notes, extract-all-notes
6. **Shape Operations**: add-shape
7. **Style Operations**: copy-text-style, copy-theme, translate-slides
8. **Export**: export-pdf, export-pptx

**Flag Variables** (package-level):
- `createPresentationFolderID`: Folder ID for new presentations
- `addSlideLayout`: Layout for new slides
- `addSlidePosition`: Position for new slides
- `styleCellBgColor`: Background color for cell styling

### internal/presentation/presentation.go
**Purpose**: Presentation-level operations

**Service Type**:
- `Service`: Wraps Slides and Drive services

**Methods**:
- `NewService(ctx, slidesService, driveService)`: Creates service instance
- `Create(ctx, title, folderID)`: Creates new presentation
- `Get(ctx, presentationID)`: Retrieves presentation by ID

### internal/slide/slide.go
**Purpose**: Slide manipulation operations

**Service Type**:
- `Service`: Wraps Slides service

**Methods**:
- `NewService(ctx, slidesService)`: Creates service instance
- `Add(ctx, presentationID, layout, position)`: Adds new slide
- `Duplicate(ctx, presentationID, slideIndex)`: Duplicates slide
- `Move(ctx, presentationID, slideIndex, newPosition)`: Moves slide
- `Remove(ctx, presentationID, slideIndex)`: Removes slide
- `Reorder(ctx, presentationID, indicesStr)`: Reorders slides

**Object ID Generation**: Uses timestamp-based unique IDs (`generateObjectID`)

### internal/table/table.go
**Purpose**: Table creation and styling

**Service Type**:
- `Service`: Wraps Slides service

**Methods**:
- `NewService(ctx, slidesService)`: Creates service instance
- `Create(ctx, presentationID, slideIndex, rows, cols)`: Creates table
- `UpdateCell(ctx, presentationID, tableID, row, col, text)`: Updates cell content
- `StyleCell(ctx, presentationID, tableID, row, col, bgColor)`: Applies cell styling

**Helper Functions**:
- `parseColor(hexColor)`: Converts hex color to OpaqueColor

### internal/text/text.go
**Purpose**: Text extraction, replacement, and search

**Service Type**:
- `Service`: Wraps Slides service

**Types**:
- `SearchResult`: Represents search result with slide index, object ID, and text

**Methods**:
- `NewService(ctx, slidesService)`: Creates service instance
- `ExtractAll(ctx, presentationID)`: Extracts all text from presentation
- `Replace(ctx, presentationID, findText, replaceText)`: Replaces text
- `Search(ctx, presentationID, query)`: Searches for text

### internal/notes/notes.go
**Purpose**: Speaker notes operations

**Service Type**:
- `Service`: Wraps Slides service

**Methods**:
- `NewService(ctx, slidesService)`: Creates service instance
- `Get(ctx, presentationID, slideIndex)`: Gets speaker notes from slide
- `Add(ctx, presentationID, slideIndex, notesContent)`: Adds speaker notes
- `ExtractAll(ctx, presentationID)`: Extracts all speaker notes

### internal/shape/shape.go
**Purpose**: Shape creation and manipulation

**Service Type**:
- `Service`: Wraps Slides service

**Methods**:
- `NewService(ctx, slidesService)`: Creates service instance
- `Add(ctx, presentationID, slideIndex, shapeType)`: Adds shape to slide

### internal/style/style.go
**Purpose**: Style and translation operations (placeholder implementations)

**Service Type**:
- `Service`: Wraps Slides service

**Methods** (all placeholders):
- `NewService(ctx, slidesService)`: Creates service instance
- `CopyTextStyle(ctx, presentationID, sourceObjectID, targetObjectID)`: Placeholder
- `CopyTheme(ctx, sourcePresentationID, targetPresentationID)`: Placeholder
- `TranslateSlides(ctx, presentationID, targetLanguage)`: Placeholder

### internal/export/export.go
**Purpose**: Export presentations to various formats

**Service Type**:
- `Service`: Wraps Drive service

**Methods**:
- `NewService(ctx, driveService)`: Creates service instance
- `ToPDF(ctx, presentationID, outputFile)`: Exports as PDF
- `ToPPTX(ctx, presentationID, outputFile)`: Exports as PowerPoint

## Key Design Patterns

### Service Pattern
Each domain package (presentation, slide, table, etc.) follows the service pattern:
1. **Service Struct**: Wraps required Google API service(s)
2. **Constructor**: `NewService(ctx, ...)` creates service instance
3. **Methods**: Domain-specific operations on the service

Example:
```go
type Service struct {
    slidesService *slides.Service
}

func NewService(ctx context.Context, slidesService *slides.Service) *Service {
    return &Service{slidesService: slidesService}
}

func (s *Service) Add(ctx context.Context, ...) error {
    // Implementation
}
```

### Separation of Concerns
- **cmd/**: Entry point only, no business logic
- **internal/auth/**: Authentication and service creation
- **internal/cli/**: Command routing and flag parsing
- **internal/{domain}/**: Business logic for specific domains
- **Context Propagation**: All service methods accept `context.Context` as first parameter

### Command Pattern (CLI)
Each operation is implemented as a cobra command with:
1. Initialization function (`initXxxCommands()`)
2. Command definition (`var xxxCmd = &cobra.Command{...}`)
3. Flag variables (package-level, grouped by command)
4. Execution function (`func runXxx(cmd, args) error`)

### Error Handling
- All errors are wrapped with context using `fmt.Errorf("message: %w", err)`
- User-facing errors written to `os.Stderr`
- Command output (IDs, JSON) written to `os.Stdout`
- Success messages use ✅ emoji

### Output Strategy
- Success messages: `fmt.Fprintf(os.Stderr, "✅ Message\n")`
- Errors: Return error from `RunE` function
- Data output: `fmt.Println(data)` or `printJSON(data)`

## API Usage Patterns

### Slides API Batch Updates
Most operations use the batch update pattern:

```go
requests := []*slides.Request{
    {
        CreateSlide: &slides.CreateSlideRequest{...},
    },
}

service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
    Requests: requests,
}).Do()
```

### Common Operations

**Get Presentation**:
```go
presentation, err := service.Presentations.Get(presentationID).Do()
```

**Create Slide**:
```go
CreateSlide: &slides.CreateSlideRequest{
    ObjectId: slideID,
    SlideLayoutReference: &slides.LayoutReference{
        PredefinedLayout: layout,
    },
}
```

**Update Cell**:
```go
InsertText: &slides.InsertTextRequest{
    ObjectId: tableID,
    CellLocation: &slides.TableCellLocation{
        RowIndex: row,
        ColumnIndex: col,
    },
    Text: text,
}
```

## Building and Development

### Build Commands
```bash
make build          # Build binary
make rebuild        # Clean all and rebuild
make install        # Install to /usr/local/bin
make uninstall      # Remove from system
make clean          # Remove build artifacts
make clean-all      # Remove build artifacts and go.mod/go.sum
make fmt            # Format code
make vet            # Run go vet
make test           # Run tests
make check          # Run fmt, vet, and test
```

### Adding New Commands

1. **Create or update service in internal/{domain}/**:
```go
package domain

type Service struct {
    slidesService *slides.Service
}

func NewService(ctx context.Context, slidesService *slides.Service) *Service {
    return &Service{slidesService: slidesService}
}

func (s *Service) DoSomething(ctx context.Context, args ...interface{}) error {
    // Implementation
    return nil
}
```

2. **Define flag variables in internal/cli/cli.go**:
```go
var (
    commandNameFlagName string
)
```

3. **Define command in internal/cli/cli.go**:
```go
var commandNameCmd = &cobra.Command{
    Use:   "command-name <args>",
    Short: "Short description",
    Args:  cobra.ExactArgs(n),
    RunE:  runCommandName,
}
```

4. **Implement execution function in internal/cli/cli.go**:
```go
func runCommandName(cmd *cobra.Command, args []string) error {
    ctx := context.Background()

    slidesService, err := auth.GetSlidesService(ctx)
    if err != nil {
        return err
    }

    svc := domain.NewService(ctx, slidesService)
    if err := svc.DoSomething(ctx, args...); err != nil {
        return err
    }

    fmt.Fprintf(os.Stderr, "✅ Success message\n")
    return nil
}
```

5. **Register command in init function in internal/cli/cli.go**:
```go
func initDomainCommands() {
    commandNameCmd.Flags().StringVar(&commandNameFlagName, "flag-name", "default", "Description")
    rootCmd.AddCommand(commandNameCmd)
}

func init() {
    // Add to existing init
    initDomainCommands()
}
```

## Known Limitations

### Placeholder Implementations
The following commands have placeholder implementations:
- `copy-text-style`: Needs full style extraction and application
- `copy-theme`: Needs master slide and layout handling
- `translate-slides`: Needs Translation API client integration

### Object ID Generation
**Fixed**: Now uses timestamp-based unique IDs via `generateObjectID(prefix)` function in each domain package:
- `slide.Add()`: Uses `generateObjectID("slide")`
- `table.Create()`: Uses `generateObjectID("table")`
- `shape.Add()`: Uses `generateObjectID("shape")`

### Context Usage
All service methods properly accept `context.Context` as first parameter. CLI commands create context with `context.Background()` for now, but this can be enhanced with cancellation support in the future.

## Testing Strategy

### Manual Testing
```bash
# Create presentation
PRES_ID=$(google-slide-manager create-presentation "Test Presentation")

# Add slide
google-slide-manager add-slide "$PRES_ID" --layout TITLE

# Export as PDF
google-slide-manager export-pdf "$PRES_ID" test.pdf
```

### Unit Testing Areas
1. Helper functions: `parseColor`, `printJSON`
2. Error handling paths
3. Command argument validation

## Security Considerations

### Credentials Storage
- Credentials stored in `~/.credentials/` directory
- Token file permissions: 0600 (owner read/write only)
- Credentials directory permissions: 0700 (owner only)

### API Scopes
Uses minimal required scopes:
- `presentations`: Read/write access to presentations
- `drive.file`: Access only to files created by this app
- `cloud-translation`: Translation API access

## Future Enhancements

### High Priority
1. Complete placeholder implementations (copy-text-style, copy-theme, translate-slides)
2. Improve object ID generation (use UUID or timestamp)
3. Add support for images and videos
4. Add bulk operations (batch create, batch export)

### Medium Priority
1. Add configuration file support
2. Add template support
3. Add progress indicators for long operations
4. Add dry-run mode for destructive operations

### Low Priority
1. Add interactive mode
2. Add presentation diff functionality
3. Add collaboration features (comments, suggestions)

## Troubleshooting for AI Agents

### Common Issues

**Authentication Errors**:
- Check if `~/.credentials/google_credentials.json` exists and is valid
- Delete `~/.credentials/google_token.json` and re-authenticate
- Verify API scopes in credentials

**Build Errors**:
- Run `make clean-all` then `make build` to regenerate dependencies
- Check Go version (requires 1.21+)
- Verify all dependencies are accessible

**API Errors**:
- Verify presentation ID is valid
- Check if slide/object indices are within range
- Ensure required APIs are enabled in Google Cloud project

**Import Issues**:
After modifying code, always run:
```bash
cd src && go mod tidy
```

## Code Modification Guidelines

### When Adding Features
1. Follow existing command pattern (command var + execution function)
2. Add flags in main.go, not in init() functions
3. Use package-level variables for flags, prefix with command name
4. Wrap all errors with context using %w
5. Write success messages to stderr, data to stdout
6. Use colored output for user-facing messages

### When Refactoring
1. Maintain alphabetical order for variables and functions where applicable
2. Keep related commands grouped together
3. Update both README.md and CLAUDE.md
4. Run `make check` before committing

### Code Style
- Follow golang standards as defined in project's golang skill
- No init() functions - use explicit initialization in main()
- Error handling: Always wrap with context, never ignore errors
- Context usage: Accept context as first parameter for service functions
- Naming: Use clear, descriptive names without unnecessary abbreviations
