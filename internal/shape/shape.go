package shape

import (
	"context"
	"fmt"
	"time"

	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides service for shape operations.
type Service struct {
	slidesService *slides.Service
}

// NewService creates a new shape service.
func NewService(ctx context.Context, slidesService *slides.Service) *Service {
	return &Service{
		slidesService: slidesService,
	}
}

// generateObjectID generates a unique object ID using timestamp.
func generateObjectID(prefix string) string {
	return fmt.Sprintf("%s_%d", prefix, time.Now().UnixNano())
}

// Add adds a shape to a slide.
func (s *Service) Add(ctx context.Context, presentationID string, slideIndex int, shapeType string) (string, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return "", fmt.Errorf("error getting presentation: %w", err)
	}

	if slideIndex >= len(presentation.Slides) {
		return "", fmt.Errorf("slide index out of range")
	}

	slideID := presentation.Slides[slideIndex].ObjectId
	shapeID := generateObjectID("shape")

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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return "", fmt.Errorf("error adding shape: %w", err)
	}

	return shapeID, nil
}
