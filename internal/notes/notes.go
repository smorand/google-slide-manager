package notes

import (
	"context"
	"fmt"
	"strings"

	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides service for notes operations.
type Service struct {
	slidesService *slides.Service
}

// NewService creates a new notes service.
func NewService(ctx context.Context, slidesService *slides.Service) *Service {
	return &Service{
		slidesService: slidesService,
	}
}

// Get retrieves speaker notes from a slide.
func (s *Service) Get(ctx context.Context, presentationID string, slideIndex int) (string, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return "", fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return "", fmt.Errorf("slide index out of range")
	}

	slide := presentation.Slides[slideIndex]
	notesPage := slide.SlideProperties.NotesPage

	if notesPage == nil {
		return "", nil
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

	return notesText.String(), nil
}

// Add adds speaker notes to a slide.
func (s *Service) Add(ctx context.Context, presentationID string, slideIndex int, notesContent string) error {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
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
				Text:           notesContent,
				InsertionIndex: 0,
			},
		},
	}

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error adding notes: %w", err)
	}

	return nil
}

// ExtractAll extracts all speaker notes from a presentation.
func (s *Service) ExtractAll(ctx context.Context, presentationID string) (map[string]string, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return nil, fmt.Errorf("error getting presentation: %w", err)
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

	return allNotes, nil
}
