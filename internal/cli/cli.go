package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"strconv"

	"github.com/spf13/cobra"

	"google-slide-manager/internal/auth"
	"google-slide-manager/internal/export"
	"google-slide-manager/internal/notes"
	"google-slide-manager/internal/presentation"
	"google-slide-manager/internal/shape"
	"google-slide-manager/internal/slide"
	"google-slide-manager/internal/style"
	"google-slide-manager/internal/table"
	"google-slide-manager/internal/text"
)

var (
	// Presentation flags
	createPresentationFolderID string

	// Slide flags
	addSlideLayout   string
	addSlidePosition int

	// Table flags
	styleCellBgColor string
)

var rootCmd = &cobra.Command{
	Use:   "google-slide-manager",
	Short: "Google Slides Manager",
	Long:  "Comprehensive Google Slides operations: create, edit, format, translate, and export presentations",
}

// Execute runs the root command.
func Execute() error {
	return rootCmd.Execute()
}

func init() {
	initPresentationCommands()
	initSlideCommands()
	initTableCommands()
	initTextCommands()
	initNotesCommands()
	initShapeCommands()
	initStyleCommands()
	initExportCommands()
}

// ==================== Presentation Commands ====================

func initPresentationCommands() {
	createPresentationCmd.Flags().StringVar(&createPresentationFolderID, "folder", "", "Folder ID to create presentation in")
	rootCmd.AddCommand(createPresentationCmd)
}

var createPresentationCmd = &cobra.Command{
	Use:   "create-presentation <title>",
	Short: "Create a new Google Slides presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runCreatePresentation,
}

func runCreatePresentation(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	title := args[0]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	driveService, err := auth.GetDriveService(ctx)
	if err != nil {
		return err
	}

	svc := presentation.NewService(ctx, slidesService, driveService)
	result, err := svc.Create(ctx, title, createPresentationFolderID)
	if err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation created: %s\n", result.Title)
	fmt.Fprintf(os.Stderr, "   ID: %s\n", result.PresentationId)
	fmt.Println(result.PresentationId)

	return nil
}

// ==================== Slide Commands ====================

func initSlideCommands() {
	addSlideCmd.Flags().StringVar(&addSlideLayout, "layout", "BLANK", "Slide layout (BLANK, TITLE, TITLE_AND_BODY, etc.)")
	addSlideCmd.Flags().IntVar(&addSlidePosition, "position", -1, "Position to insert slide (-1 for end)")
	rootCmd.AddCommand(addSlideCmd)
	rootCmd.AddCommand(duplicateSlideCmd)
	rootCmd.AddCommand(removeSlideCmd)
	rootCmd.AddCommand(moveSlideCmd)
	rootCmd.AddCommand(reorderSlidesCmd)
}

var addSlideCmd = &cobra.Command{
	Use:   "add-slide <presentation-id>",
	Short: "Add a new slide to presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runAddSlide,
}

func runAddSlide(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := slide.NewService(ctx, slidesService)
	slideID, err := svc.Add(ctx, presentationID, addSlideLayout, addSlidePosition)
	if err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slide added with layout %s\n", addSlideLayout)
	fmt.Println(slideID)

	return nil
}

var duplicateSlideCmd = &cobra.Command{
	Use:   "duplicate-slide <presentation-id> <slide-index>",
	Short: "Duplicate an existing slide",
	Args:  cobra.ExactArgs(2),
	RunE:  runDuplicateSlide,
}

func runDuplicateSlide(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := slide.NewService(ctx, slidesService)
	if err := svc.Duplicate(ctx, presentationID, slideIndex); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slide duplicated\n")
	return nil
}

var removeSlideCmd = &cobra.Command{
	Use:   "remove-slide <presentation-id> <slide-index>",
	Short: "Remove a slide from presentation",
	Args:  cobra.ExactArgs(2),
	RunE:  runRemoveSlide,
}

func runRemoveSlide(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := slide.NewService(ctx, slidesService)
	if err := svc.Remove(ctx, presentationID, slideIndex); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slide removed\n")
	return nil
}

var moveSlideCmd = &cobra.Command{
	Use:   "move-slide <presentation-id> <slide-index> <new-position>",
	Short: "Move a slide to new position",
	Args:  cobra.ExactArgs(3),
	RunE:  runMoveSlide,
}

func runMoveSlide(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	newPosition, err := strconv.Atoi(args[2])
	if err != nil {
		return fmt.Errorf("invalid new position: %w", err)
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := slide.NewService(ctx, slidesService)
	if err := svc.Move(ctx, presentationID, slideIndex, newPosition); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slide moved to position %d\n", newPosition)
	return nil
}

var reorderSlidesCmd = &cobra.Command{
	Use:   "reorder-slides <presentation-id> <indices>",
	Short: "Reorder slides (comma-separated indices)",
	Args:  cobra.ExactArgs(2),
	RunE:  runReorderSlides,
}

func runReorderSlides(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	indicesStr := args[1]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := slide.NewService(ctx, slidesService)
	if err := svc.Reorder(ctx, presentationID, indicesStr); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slides reordered\n")
	return nil
}

// ==================== Table Commands ====================

func initTableCommands() {
	styleCellCmd.Flags().StringVar(&styleCellBgColor, "bg-color", "", "Background color (hex, e.g., #FF0000)")
	rootCmd.AddCommand(createTableCmd)
	rootCmd.AddCommand(updateCellCmd)
	rootCmd.AddCommand(styleCellCmd)
}

var createTableCmd = &cobra.Command{
	Use:   "create-table <presentation-id> <slide-index> <rows> <cols>",
	Short: "Create a table on a slide",
	Args:  cobra.ExactArgs(4),
	RunE:  runCreateTable,
}

func runCreateTable(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	rows, err := strconv.ParseInt(args[2], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid rows: %w", err)
	}

	cols, err := strconv.ParseInt(args[3], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid cols: %w", err)
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := table.NewService(ctx, slidesService)
	tableID, err := svc.Create(ctx, presentationID, slideIndex, rows, cols)
	if err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Table created (%dx%d)\n", rows, cols)
	fmt.Println(tableID)

	return nil
}

var updateCellCmd = &cobra.Command{
	Use:   "update-cell <presentation-id> <table-id> <row> <col> <text>",
	Short: "Update table cell content",
	Args:  cobra.ExactArgs(5),
	RunE:  runUpdateCell,
}

func runUpdateCell(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	tableID := args[1]

	row, err := strconv.ParseInt(args[2], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid row: %w", err)
	}

	col, err := strconv.ParseInt(args[3], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid col: %w", err)
	}

	textContent := args[4]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := table.NewService(ctx, slidesService)
	if err := svc.UpdateCell(ctx, presentationID, tableID, row, col, textContent); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Cell updated (row %d, col %d)\n", row, col)
	return nil
}

var styleCellCmd = &cobra.Command{
	Use:   "style-cell <presentation-id> <table-id> <row> <col>",
	Short: "Style table cell (background color)",
	Args:  cobra.ExactArgs(4),
	RunE:  runStyleCell,
}

func runStyleCell(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	tableID := args[1]

	row, err := strconv.ParseInt(args[2], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid row: %w", err)
	}

	col, err := strconv.ParseInt(args[3], 10, 64)
	if err != nil {
		return fmt.Errorf("invalid col: %w", err)
	}

	if styleCellBgColor == "" {
		return fmt.Errorf("background color is required (--bg-color)")
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := table.NewService(ctx, slidesService)
	if err := svc.StyleCell(ctx, presentationID, tableID, row, col, styleCellBgColor); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Cell styled (row %d, col %d)\n", row, col)
	return nil
}

// ==================== Text Commands ====================

func initTextCommands() {
	rootCmd.AddCommand(replaceTextCmd)
	rootCmd.AddCommand(extractAllTextCmd)
	rootCmd.AddCommand(searchTextCmd)
}

var replaceTextCmd = &cobra.Command{
	Use:   "replace-text <presentation-id> <find> <replace>",
	Short: "Find and replace text in presentation",
	Args:  cobra.ExactArgs(3),
	RunE:  runReplaceText,
}

func runReplaceText(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	findText := args[1]
	replaceText := args[2]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := text.NewService(ctx, slidesService)
	if err := svc.Replace(ctx, presentationID, findText, replaceText); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Text replaced: '%s' -> '%s'\n", findText, replaceText)
	return nil
}

var extractAllTextCmd = &cobra.Command{
	Use:   "extract-all-text <presentation-id>",
	Short: "Extract all text from presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runExtractAllText,
}

func runExtractAllText(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := text.NewService(ctx, slidesService)
	allText, err := svc.ExtractAll(ctx, presentationID)
	if err != nil {
		return err
	}

	fmt.Println(allText)
	return nil
}

var searchTextCmd = &cobra.Command{
	Use:   "search-text <presentation-id> <query>",
	Short: "Search for text in presentation",
	Args:  cobra.ExactArgs(2),
	RunE:  runSearchText,
}

func runSearchText(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	query := args[1]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := text.NewService(ctx, slidesService)
	results, err := svc.Search(ctx, presentationID, query)
	if err != nil {
		return err
	}

	return printJSON(results)
}

// ==================== Notes Commands ====================

func initNotesCommands() {
	rootCmd.AddCommand(getNotesCmd)
	rootCmd.AddCommand(addNotesCmd)
	rootCmd.AddCommand(extractAllNotesCmd)
}

var getNotesCmd = &cobra.Command{
	Use:   "get-notes <presentation-id> <slide-index>",
	Short: "Get speaker notes from a slide",
	Args:  cobra.ExactArgs(2),
	RunE:  runGetNotes,
}

func runGetNotes(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := notes.NewService(ctx, slidesService)
	notesText, err := svc.Get(ctx, presentationID, slideIndex)
	if err != nil {
		return err
	}

	fmt.Println(notesText)
	return nil
}

var addNotesCmd = &cobra.Command{
	Use:   "add-notes <presentation-id> <slide-index> <notes>",
	Short: "Add speaker notes to a slide",
	Args:  cobra.ExactArgs(3),
	RunE:  runAddNotes,
}

func runAddNotes(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	notesContent := args[2]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := notes.NewService(ctx, slidesService)
	if err := svc.Add(ctx, presentationID, slideIndex, notesContent); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Notes added to slide %d\n", slideIndex)
	return nil
}

var extractAllNotesCmd = &cobra.Command{
	Use:   "extract-all-notes <presentation-id>",
	Short: "Extract all speaker notes from presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runExtractAllNotes,
}

func runExtractAllNotes(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := notes.NewService(ctx, slidesService)
	allNotes, err := svc.ExtractAll(ctx, presentationID)
	if err != nil {
		return err
	}

	return printJSON(allNotes)
}

// ==================== Shape Commands ====================

func initShapeCommands() {
	rootCmd.AddCommand(addShapeCmd)
}

var addShapeCmd = &cobra.Command{
	Use:   "add-shape <presentation-id> <slide-index> <shape-type>",
	Short: "Add a shape to a slide (RECTANGLE, ELLIPSE, etc.)",
	Args:  cobra.ExactArgs(3),
	RunE:  runAddShape,
}

func runAddShape(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	slideIndex, err := strconv.Atoi(args[1])
	if err != nil {
		return fmt.Errorf("invalid slide index: %w", err)
	}

	shapeType := args[2]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := shape.NewService(ctx, slidesService)
	shapeID, err := svc.Add(ctx, presentationID, slideIndex, shapeType)
	if err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Shape added: %s\n", shapeType)
	fmt.Println(shapeID)

	return nil
}

// ==================== Style Commands ====================

func initStyleCommands() {
	rootCmd.AddCommand(copyTextStyleCmd)
	rootCmd.AddCommand(copyThemeCmd)
	rootCmd.AddCommand(translateSlidesCmd)
}

var copyTextStyleCmd = &cobra.Command{
	Use:   "copy-text-style <presentation-id> <source-object-id> <target-object-id>",
	Short: "Copy text style from one element to another",
	Args:  cobra.ExactArgs(3),
	RunE:  runCopyTextStyle,
}

func runCopyTextStyle(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	sourceObjectID := args[1]
	targetObjectID := args[2]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := style.NewService(ctx, slidesService)
	if err := svc.CopyTextStyle(ctx, presentationID, sourceObjectID, targetObjectID); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Text style copied\n")
	return nil
}

var copyThemeCmd = &cobra.Command{
	Use:   "copy-theme <source-presentation-id> <target-presentation-id>",
	Short: "Copy theme from one presentation to another",
	Args:  cobra.ExactArgs(2),
	RunE:  runCopyTheme,
}

func runCopyTheme(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	sourcePresentationID := args[0]
	targetPresentationID := args[1]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := style.NewService(ctx, slidesService)
	if err := svc.CopyTheme(ctx, sourcePresentationID, targetPresentationID); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Theme copied\n")
	return nil
}

var translateSlidesCmd = &cobra.Command{
	Use:   "translate-slides <presentation-id> <target-language>",
	Short: "Translate slides to target language (e.g., fr, es, de)",
	Args:  cobra.ExactArgs(2),
	RunE:  runTranslateSlides,
}

func runTranslateSlides(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	targetLanguage := args[1]

	slidesService, err := auth.GetSlidesService(ctx)
	if err != nil {
		return err
	}

	svc := style.NewService(ctx, slidesService)
	if err := svc.TranslateSlides(ctx, presentationID, targetLanguage); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Slides translated\n")
	return nil
}

// ==================== Export Commands ====================

func initExportCommands() {
	rootCmd.AddCommand(exportPdfCmd)
	rootCmd.AddCommand(exportPptxCmd)
}

var exportPdfCmd = &cobra.Command{
	Use:   "export-pdf <presentation-id> <output-file>",
	Short: "Export presentation as PDF",
	Args:  cobra.ExactArgs(2),
	RunE:  runExportPdf,
}

func runExportPdf(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	outputFile := args[1]

	driveService, err := auth.GetDriveService(ctx)
	if err != nil {
		return err
	}

	svc := export.NewService(ctx, driveService)
	if err := svc.ToPDF(ctx, presentationID, outputFile); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation exported as PDF: %s\n", outputFile)
	return nil
}

var exportPptxCmd = &cobra.Command{
	Use:   "export-pptx <presentation-id> <output-file>",
	Short: "Export presentation as PowerPoint",
	Args:  cobra.ExactArgs(2),
	RunE:  runExportPptx,
}

func runExportPptx(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]
	outputFile := args[1]

	driveService, err := auth.GetDriveService(ctx)
	if err != nil {
		return err
	}

	svc := export.NewService(ctx, driveService)
	if err := svc.ToPPTX(ctx, presentationID, outputFile); err != nil {
		return err
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation exported as PPTX: %s\n", outputFile)
	return nil
}

// ==================== Helper Functions ====================

func printJSON(v interface{}) error {
	data, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		return err
	}
	fmt.Println(string(data))
	return nil
}
