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

```
google-slide-manager/
├── Makefile              # Build automation
├── README.md             # Human-oriented documentation
├── CLAUDE.md             # This file - AI-oriented documentation
└── src/                  # Go source code
    ├── main.go           # Entry point, command registration, flag initialization
    ├── cli.go            # All CLI command definitions and implementations
    └── auth.go           # OAuth2 authentication and credential management
```

## File Descriptions

### src/main.go
**Purpose**: Application entry point and command registration

**Key Components**:
- `rootCmd`: Root cobra command definition
- `main()`: Initializes flags and registers all commands

**Important**: Flags are initialized in `main()` function, not in `init()` functions (per golang standards).

### src/cli.go
**Purpose**: All CLI command implementations

**Structure**:
- Package-level variables for command flags (grouped by command)
- Command definitions (cobra.Command structs)
- Command execution functions (runXxx functions)
- Helper functions (parseColor, printJSON)

**Command Categories**:
1. **Presentation Operations**: create-presentation
2. **Slide Operations**: add-slide, duplicate-slide, remove-slide, move-slide, reorder-slides
3. **Table Operations**: create-table, update-cell, style-cell
4. **Text Operations**: replace-text, extract-all-text, search-text
5. **Notes Operations**: get-notes, add-notes, extract-all-notes
6. **Shape Operations**: add-shape
7. **Style Operations**: copy-text-style, copy-theme
8. **Translation**: translate-slides
9. **Export**: export-pdf, export-pptx

**Flag Variables** (package-level):
- `createPresentationFolderID`: Folder ID for new presentations
- `addSlideLayout`: Layout for new slides
- `addSlidePosition`: Position for new slides
- `styleCellBgColor`: Background color for cell styling

### src/auth.go
**Purpose**: Google OAuth2 authentication and credential management

**Key Functions**:
- `getCredentialsPath()`: Returns path to `~/.credentials` directory
- `getClient(ctx)`: Creates authenticated HTTP client
- `getSlidesService(ctx)`: Creates Slides API service
- `getDriveService(ctx)`: Creates Drive API service
- `getTokenFromWeb(config)`: Performs OAuth2 web flow
- `tokenFromFile(path)`: Loads existing token
- `saveToken(path, token)`: Saves OAuth2 token

**Credentials Location**:
- Credentials: `~/.credentials/google_credentials.json`
- Token: `~/.credentials/google_token.json`

**API Scopes**:
```go
scopes = []string{
    slides.PresentationsScope,
    drive.DriveFileScope,
    "https://www.googleapis.com/auth/cloud-translation",
}
```

## Key Design Patterns

### Command Pattern
Each operation is implemented as a separate cobra command with:
1. Command definition (`var xxxCmd = &cobra.Command{...}`)
2. Flag variables (package-level, grouped by command)
3. Execution function (`func runXxx(cmd, args) error`)

### Error Handling
- All errors are wrapped with context using `fmt.Errorf("message: %w", err)`
- User-facing errors written to `os.Stderr`
- Command output (IDs, JSON) written to `os.Stdout`
- Success messages use colored output (green checkmarks)

### Output Strategy
- Success messages: `fmt.Fprintf(os.Stderr, "%s✅ Message%s\n", green, green)`
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

1. **Define flag variables** (in cli.go):
```go
var (
    commandNameFlagName string
)
```

2. **Define command** (in cli.go):
```go
var commandNameCmd = &cobra.Command{
    Use:   "command-name <args>",
    Short: "Short description",
    Args:  cobra.ExactArgs(n),
    RunE:  runCommandName,
}
```

3. **Implement execution function** (in cli.go):
```go
func runCommandName(cmd *cobra.Command, args []string) error {
    ctx := context.Background()
    // Implementation
    return nil
}
```

4. **Register command and flags** (in main.go):
```go
func main() {
    // Initialize flags
    commandNameCmd.Flags().StringVar(&commandNameFlagName, "flag-name", "default", "Description")

    // Add command
    rootCmd.AddCommand(commandNameCmd)

    // ... rest of main
}
```

## Known Limitations

### Placeholder Implementations
The following commands have placeholder implementations:
- `copy-text-style`: Needs full style extraction and application
- `copy-theme`: Needs master slide and layout handling
- `translate-slides`: Needs Translation API client integration

### Object ID Generation
Some commands use simple object ID generation:
- `add-slide`: Uses `fmt.Sprintf("slide_%d", len(args))` - should use timestamp or UUID
- `create-table`: Uses `fmt.Sprintf("table_%d", slideIndex)` - should use timestamp or UUID
- `add-shape`: Uses `fmt.Sprintf("shape_%d", slideIndex)` - should use timestamp or UUID

### Context Usage
Currently all commands use `context.Background()` directly in the execution function. For better cancellation support, consider accepting context as parameter.

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
