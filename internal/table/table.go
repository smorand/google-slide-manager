package table

import (
	"context"
	"fmt"
	"strings"
	"time"

	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides service for table operations.
type Service struct {
	slidesService *slides.Service
}

// NewService creates a new table service.
func NewService(ctx context.Context, slidesService *slides.Service) *Service {
	return &Service{
		slidesService: slidesService,
	}
}

// generateObjectID generates a unique object ID using timestamp.
func generateObjectID(prefix string) string {
	return fmt.Sprintf("%s_%d", prefix, time.Now().UnixNano())
}

// Create creates a table on a slide.
func (s *Service) Create(ctx context.Context, presentationID string, slideIndex int, rows int64, cols int64) (string, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return "", fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return "", fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId
	tableID := generateObjectID("table")

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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return "", fmt.Errorf("error creating table: %w", err)
	}

	return tableID, nil
}

// UpdateCell updates a table cell content.
func (s *Service) UpdateCell(ctx context.Context, presentationID string, tableID string, row int64, col int64, text string) error {
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

	_, err := s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error updating cell: %w", err)
	}

	return nil
}

// parseColor converts hex color to OpaqueColor.
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

// StyleCell applies styling to a table cell.
func (s *Service) StyleCell(ctx context.Context, presentationID string, tableID string, row int64, col int64, bgColor string) error {
	requests := []*slides.Request{
		{
			UpdateTableCellProperties: &slides.UpdateTableCellPropertiesRequest{
				ObjectId: tableID,
				TableCellProperties: &slides.TableCellProperties{
					TableCellBackgroundFill: &slides.TableCellBackgroundFill{
						SolidFill: &slides.SolidFill{
							Color: parseColor(bgColor),
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

	_, err := s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error styling cell: %w", err)
	}

	return nil
}
