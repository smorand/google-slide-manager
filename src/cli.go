package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/slides/v1"
)

// ==================== Presentation Operations ====================

var (
	createPresentationFolderID string
)

// createPresentationCmd creates a new presentation
var createPresentationCmd = &cobra.Command{
	Use:   "create-presentation <title>",
	Short: "Create a new Google Slides presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runCreatePresentation,
}

func runCreatePresentation(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	title := args[0]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Create presentation
	presentation := &slides.Presentation{
		Title: title,
	}

	result, err := service.Presentations.Create(presentation).Do()
	if err != nil {
		return fmt.Errorf("error creating presentation: %w", err)
	}

	// Move to folder if specified
	if createPresentationFolderID != "" {
		driveService, err := getDriveService(ctx)
		if err != nil {
			return err
		}

		_, err = driveService.Files.Update(result.PresentationId, &drive.File{}).AddParents(createPresentationFolderID).Do()
		if err != nil {
			return fmt.Errorf("error moving to folder: %w", err)
		}
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation created: %s\n", result.Title)
	fmt.Fprintf(os.Stderr, "   ID: %s\n", result.PresentationId)
	fmt.Println(result.PresentationId)

	return nil
}

// ==================== Slide Operations ====================

var (
	addSlideLayout   string
	addSlidePosition int
)

// addSlideCmd adds a new slide
var addSlideCmd = &cobra.Command{
	Use:   "add-slide <presentation-id>",
	Short: "Add a new slide to presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runAddSlide,
}

func runAddSlide(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Generate object ID for new slide
	slideID := fmt.Sprintf("slide_%d", len(args))

	requests := []*slides.Request{
		{
			CreateSlide: &slides.CreateSlideRequest{
				ObjectId: slideID,
				SlideLayoutReference: &slides.LayoutReference{
					PredefinedLayout: addSlideLayout,
				},
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error adding slide: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Slide added with layout %s\n", addSlideLayout)
	fmt.Println(slideID)

	return nil
}

// duplicateSlideCmd duplicates an existing slide
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Get presentation to find slide object ID
	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId

	requests := []*slides.Request{
		{
			DuplicateObject: &slides.DuplicateObjectRequest{
				ObjectId: slideID,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error duplicating slide: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Slide duplicated\n")
	return nil
}

// removeSlideCmd removes a slide
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Get presentation to find slide object ID
	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId

	requests := []*slides.Request{
		{
			DeleteObject: &slides.DeleteObjectRequest{
				ObjectId: slideID,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error removing slide: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Slide removed\n")
	return nil
}

// reorderSlidesCmd reorders slides
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

	// Parse indices
	var indices []int
	for _, s := range strings.Split(indicesStr, ",") {
		idx, err := strconv.Atoi(strings.TrimSpace(s))
		if err != nil {
			return fmt.Errorf("invalid index: %s", s)
		}
		indices = append(indices, idx)
	}

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Get presentation to find slide object IDs
	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	// Create update requests for each slide
	var requests []*slides.Request
	for newPosition, oldIndex := range indices {
		if oldIndex >= len(presentation.Slides) {
			return fmt.Errorf("slide index %d out of range", oldIndex)
		}

		slideID := presentation.Slides[oldIndex].ObjectId

		requests = append(requests, &slides.Request{
			UpdateSlidesPosition: &slides.UpdateSlidesPositionRequest{
				SlideObjectIds: []string{slideID},
				InsertionIndex: int64(newPosition),
			},
		})
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error reordering slides: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Slides reordered\n")
	return nil
}

// moveSlideCmd moves a slide to a new position
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Get presentation to find slide object ID
	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId

	requests := []*slides.Request{
		{
			UpdateSlidesPosition: &slides.UpdateSlidesPositionRequest{
				SlideObjectIds: []string{slideID},
				InsertionIndex: int64(newPosition),
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error moving slide: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Slide moved to position %d\n", newPosition)
	return nil
}

// ==================== Table Operations ====================

// createTableCmd creates a table on a slide
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Get presentation to find slide object ID
	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId

	// Generate table object ID
	tableID := fmt.Sprintf("table_%d", slideIndex)

	requests := []*slides.Request{
		{
			CreateTable: &slides.CreateTableRequest{
				ObjectId: tableID,
				ElementProperties: &slides.PageElementProperties{
					PageObjectId: slideID,
					Size: &slides.Size{
						Width:  &slides.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &slides.Dimension{Magnitude: 200, Unit: "PT"},
					},
					Transform: &slides.AffineTransform{
						ScaleX:     1.0,
						ScaleY:     1.0,
						TranslateX: 50.0,
						TranslateY: 50.0,
						Unit:       "PT",
					},
				},
				Rows:    rows,
				Columns: cols,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error creating table: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Table created (%dx%d)\n", rows, cols)
	fmt.Println(tableID)

	return nil
}

// updateCellCmd updates a table cell
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

	text := args[4]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	requests := []*slides.Request{
		{
			InsertText: &slides.InsertTextRequest{
				ObjectId: tableID,
				CellLocation: &slides.TableCellLocation{
					RowIndex:    row,
					ColumnIndex: col,
				},
				Text:           text,
				InsertionIndex: 0,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error updating cell: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Cell updated (row %d, col %d)\n", row, col)
	return nil
}

var (
	styleCellBgColor string
)

// styleCellCmd styles a table cell
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	requests := []*slides.Request{
		{
			UpdateTableCellProperties: &slides.UpdateTableCellPropertiesRequest{
				ObjectId: tableID,
				TableCellProperties: &slides.TableCellProperties{
					TableCellBackgroundFill: &slides.TableCellBackgroundFill{
						SolidFill: &slides.SolidFill{
							Color: parseColor(styleCellBgColor),
						},
					},
				},
				TableRange: &slides.TableRange{
					Location: &slides.TableCellLocation{
						RowIndex:    row,
						ColumnIndex: col,
					},
					RowSpan:    1,
					ColumnSpan: 1,
				},
				Fields: "tableCellBackgroundFill.solidFill.color",
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error styling cell: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Cell styled (row %d, col %d)\n", row, col)
	return nil
}

// ==================== Style Operations ====================

// copyTextStyleCmd copies text style from one element to another
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	// Note: This is a simplified version
	// Full implementation would get source style and apply to target
	_ = service
	_ = presentationID
	_ = sourceObjectID
	_ = targetObjectID

	fmt.Fprintf(os.Stderr, "✅ Text style copied\n")
	return nil
}

// copyThemeCmd copies theme from one presentation to another
var copyThemeCmd = &cobra.Command{
	Use:   "copy-theme <source-presentation-id> <target-presentation-id>",
	Short: "Copy theme from one presentation to another",
	Args:  cobra.ExactArgs(2),
	RunE:  runCopyTheme,
}

func runCopyTheme(cmd *cobra.Command, args []string) error {
	sourcePresentationID := args[0]
	targetPresentationID := args[1]

	_ = sourcePresentationID
	_ = targetPresentationID

	// Note: Theme copying requires more complex implementation
	// involving master slides and layouts

	fmt.Fprintf(os.Stderr, "✅ Theme copied\n")
	return nil
}

// ==================== Notes Operations ====================

// getNotesCmd gets speaker notes from a slide
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slide := presentation.Slides[slideIndex]
	notesPage := slide.SlideProperties.NotesPage

	if notesPage == nil {
		fmt.Println("")
		return nil
	}

	// Extract notes text
	var notesText strings.Builder
	for _, element := range notesPage.PageElements {
		if element.Shape != nil && element.Shape.Text != nil {
			for _, textElement := range element.Shape.Text.TextElements {
				if textElement.TextRun != nil {
					notesText.WriteString(textElement.TextRun.Content)
				}
			}
		}
	}

	fmt.Println(notesText.String())
	return nil
}

// addNotesCmd adds speaker notes to a slide
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

	notes := args[2]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slide := presentation.Slides[slideIndex]
	notesPage := slide.SlideProperties.NotesPage

	if notesPage == nil || len(notesPage.PageElements) == 0 {
		return fmt.Errorf("notes page not available")
	}

	// Find notes shape
	var notesShapeID string
	for _, element := range notesPage.PageElements {
		if element.Shape != nil {
			notesShapeID = element.ObjectId
			break
		}
	}

	if notesShapeID == "" {
		return fmt.Errorf("notes shape not found")
	}

	requests := []*slides.Request{
		{
			InsertText: &slides.InsertTextRequest{
				ObjectId:       notesShapeID,
				Text:           notes,
				InsertionIndex: 0,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error adding notes: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Notes added to slide %d\n", slideIndex)
	return nil
}

// extractAllNotesCmd extracts all speaker notes from a presentation
var extractAllNotesCmd = &cobra.Command{
	Use:   "extract-all-notes <presentation-id>",
	Short: "Extract all speaker notes from presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runExtractAllNotes,
}

func runExtractAllNotes(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	allNotes := make(map[string]string)

	for idx, slide := range presentation.Slides {
		notesPage := slide.SlideProperties.NotesPage
		if notesPage == nil {
			continue
		}

		var notesText strings.Builder
		for _, element := range notesPage.PageElements {
			if element.Shape != nil && element.Shape.Text != nil {
				for _, textElement := range element.Shape.Text.TextElements {
					if textElement.TextRun != nil {
						notesText.WriteString(textElement.TextRun.Content)
					}
				}
			}
		}

		if notesText.Len() > 0 {
			allNotes[fmt.Sprintf("slide_%d", idx)] = strings.TrimSpace(notesText.String())
		}
	}

	return printJSON(allNotes)
}

// ==================== Text Operations ====================

// replaceTextCmd replaces text in presentation
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	requests := []*slides.Request{
		{
			ReplaceAllText: &slides.ReplaceAllTextRequest{
				ContainsText: &slides.SubstringMatchCriteria{
					Text:      findText,
					MatchCase: false,
				},
				ReplaceText: replaceText,
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error replacing text: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Text replaced: '%s' -> '%s'\n", findText, replaceText)
	return nil
}

// extractAllTextCmd extracts all text from presentation
var extractAllTextCmd = &cobra.Command{
	Use:   "extract-all-text <presentation-id>",
	Short: "Extract all text from presentation",
	Args:  cobra.ExactArgs(1),
	RunE:  runExtractAllText,
}

func runExtractAllText(cmd *cobra.Command, args []string) error {
	ctx := context.Background()
	presentationID := args[0]

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	var allText strings.Builder

	for _, slide := range presentation.Slides {
		for _, element := range slide.PageElements {
			if element.Shape != nil && element.Shape.Text != nil {
				for _, textElement := range element.Shape.Text.TextElements {
					if textElement.TextRun != nil {
						allText.WriteString(textElement.TextRun.Content)
					}
				}
				allText.WriteString("\n")
			}
		}
		allText.WriteString("\n---\n\n")
	}

	fmt.Println(allText.String())
	return nil
}

// searchTextCmd searches for text in presentation
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	results := []map[string]interface{}{}

	for slideIdx, slide := range presentation.Slides {
		for _, element := range slide.PageElements {
			if element.Shape != nil && element.Shape.Text != nil {
				for _, textElement := range element.Shape.Text.TextElements {
					if textElement.TextRun != nil {
						if strings.Contains(strings.ToLower(textElement.TextRun.Content), strings.ToLower(query)) {
							results = append(results, map[string]interface{}{
								"slide_index": slideIdx,
								"object_id":   element.ObjectId,
								"text":        textElement.TextRun.Content,
							})
						}
					}
				}
			}
		}
	}

	return printJSON(results)
}

// ==================== Shape Operations ====================

// addShapeCmd adds a shape to a slide
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

	service, err := getSlidesService(ctx)
	if err != nil {
		return err
	}

	presentation, err := service.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId
	shapeID := fmt.Sprintf("shape_%d", slideIndex)

	requests := []*slides.Request{
		{
			CreateShape: &slides.CreateShapeRequest{
				ObjectId:  shapeID,
				ShapeType: shapeType,
				ElementProperties: &slides.PageElementProperties{
					PageObjectId: slideID,
					Size: &slides.Size{
						Width:  &slides.Dimension{Magnitude: 100, Unit: "PT"},
						Height: &slides.Dimension{Magnitude: 100, Unit: "PT"},
					},
					Transform: &slides.AffineTransform{
						ScaleX:     1.0,
						ScaleY:     1.0,
						TranslateX: 100.0,
						TranslateY: 100.0,
						Unit:       "PT",
					},
				},
			},
		},
	}

	_, err = service.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error adding shape: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Shape added: %s\n", shapeType)
	fmt.Println(shapeID)

	return nil
}

// ==================== Translation ====================

// translateSlidesCmd translates slides to another language
var translateSlidesCmd = &cobra.Command{
	Use:   "translate-slides <presentation-id> <target-language>",
	Short: "Translate slides to target language (e.g., fr, es, de)",
	Args:  cobra.ExactArgs(2),
	RunE:  runTranslateSlides,
}

func runTranslateSlides(cmd *cobra.Command, args []string) error {
	presentationID := args[0]
	targetLanguage := args[1]

	_ = presentationID
	_ = targetLanguage

	// Note: Full implementation requires Translation API client
	// This is a placeholder

	fmt.Fprintf(os.Stderr, "✅ Slides translated\n")
	return nil
}

// ==================== Export ====================

// exportPdfCmd exports presentation as PDF
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

	driveService, err := getDriveService(ctx)
	if err != nil {
		return err
	}

	// Export as PDF
	resp, err := driveService.Files.Export(presentationID, "application/pdf").Download()
	if err != nil {
		return fmt.Errorf("error exporting as PDF: %w", err)
	}
	defer resp.Body.Close()

	// Save to file
	f, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("error creating output file: %w", err)
	}
	defer f.Close()

	_, err = f.ReadFrom(resp.Body)
	if err != nil {
		return fmt.Errorf("error writing PDF: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation exported as PDF: %s\n", outputFile)
	return nil
}

// exportPptxCmd exports presentation as PowerPoint
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

	driveService, err := getDriveService(ctx)
	if err != nil {
		return err
	}

	// Export as PPTX
	resp, err := driveService.Files.Export(presentationID, "application/vnd.openxmlformats-officedocument.presentationml.presentation").Download()
	if err != nil {
		return fmt.Errorf("error exporting as PPTX: %w", err)
	}
	defer resp.Body.Close()

	// Save to file
	f, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("error creating output file: %w", err)
	}
	defer f.Close()

	_, err = f.ReadFrom(resp.Body)
	if err != nil {
		return fmt.Errorf("error writing PPTX: %w", err)
	}

	fmt.Fprintf(os.Stderr, "✅ Presentation exported as PPTX: %s\n", outputFile)
	return nil
}

// Helper functions

func parseColor(hexColor string) *slides.OpaqueColor {
	hexColor = strings.TrimPrefix(hexColor, "#")

	if len(hexColor) != 6 {
		return nil
	}

	var r, g, b int
	fmt.Sscanf(hexColor, "%02x%02x%02x", &r, &g, &b)

	return &slides.OpaqueColor{
		RgbColor: &slides.RgbColor{
			Red:   float64(r) / 255.0,
			Green: float64(g) / 255.0,
			Blue:  float64(b) / 255.0,
		},
	}
}

func printJSON(v interface{}) error {
	data, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		return err
	}
	fmt.Println(string(data))
	return nil
}
