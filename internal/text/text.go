package text

import (
	"context"
	"fmt"
	"strings"

	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides service for text operations.
type Service struct {
	slidesService *slides.Service
}

// SearchResult represents a text search result.
type SearchResult struct {
	SlideIndex int    `json:"slide_index"`
	ObjectID   string `json:"object_id"`
	Text       string `json:"text"`
}

// NewService creates a new text service.
func NewService(ctx context.Context, slidesService *slides.Service) *Service {
	return &Service{
		slidesService: slidesService,
	}
}

// ExtractAll extracts all text from a presentation.
func (s *Service) ExtractAll(ctx context.Context, presentationID string) (string, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return "", fmt.Errorf("error getting presentation: %w", err)
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

	return allText.String(), nil
}

// Replace replaces all occurrences of find text with replace text.
func (s *Service) Replace(ctx context.Context, presentationID string, findText string, replaceText string) error {
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

	_, err := s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error replacing text: %w", err)
	}

	return nil
}

// Search searches for text in a presentation and returns matches.
func (s *Service) Search(ctx context.Context, presentationID string, query string) ([]SearchResult, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return nil, fmt.Errorf("error getting presentation: %w", err)
	}

	var results []SearchResult

	for slideIdx, slide := range presentation.Slides {
		for _, element := range slide.PageElements {
			if element.Shape != nil && element.Shape.Text != nil {
				for _, textElement := range element.Shape.Text.TextElements {
					if textElement.TextRun != nil {
						if strings.Contains(strings.ToLower(textElement.TextRun.Content), strings.ToLower(query)) {
							results = append(results, SearchResult{
								SlideIndex: slideIdx,
								ObjectID:   element.ObjectId,
								Text:       textElement.TextRun.Content,
							})
						}
					}
				}
			}
		}
	}

	return results, nil
}
