package slide

import (
	"context"
	"fmt"
	"strconv"
	"strings"
	"time"

	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides service for slide operations.
type Service struct {
	slidesService *slides.Service
}

// NewService creates a new slide service.
func NewService(ctx context.Context, slidesService *slides.Service) *Service {
	return &Service{
		slidesService: slidesService,
	}
}

// generateObjectID generates a unique object ID using timestamp.
func generateObjectID(prefix string) string {
	return fmt.Sprintf("%s_%d", prefix, time.Now().UnixNano())
}

// Add adds a new slide to the presentation.
func (s *Service) Add(ctx context.Context, presentationID string, layout string, position int) (string, error) {
	slideID := generateObjectID("slide")

	requests := []*slides.Request{
		{
			CreateSlide: &slides.CreateSlideRequest{
				ObjectId: slideID,
				SlideLayoutReference: &slides.LayoutReference{
					PredefinedLayout: layout,
				},
			},
		},
	}

	if position >= 0 {
		requests[0].CreateSlide.InsertionIndex = int64(position)
	}

	_, err := s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return "", fmt.Errorf("error adding slide: %w", err)
	}

	return slideID, nil
}

// Duplicate duplicates an existing slide.
func (s *Service) Duplicate(ctx context.Context, presentationID string, slideIndex int) error {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error duplicating slide: %w", err)
	}

	return nil
}

// Move moves a slide to a new position.
func (s *Service) Move(ctx context.Context, presentationID string, slideIndex int, newPosition int) error {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error moving slide: %w", err)
	}

	return nil
}

// Remove removes a slide from the presentation.
func (s *Service) Remove(ctx context.Context, presentationID string, slideIndex int) error {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error removing slide: %w", err)
	}

	return nil
}

// Reorder reorders slides according to the provided indices.
func (s *Service) Reorder(ctx context.Context, presentationID string, indicesStr string) error {
	var indices []int
	for _, s := range strings.Split(indicesStr, ",") {
		idx, err := strconv.Atoi(strings.TrimSpace(s))
		if err != nil {
			return fmt.Errorf("invalid index: %s", s)
		}
		indices = append(indices, idx)
	}

	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return fmt.Errorf("error getting presentation: %w", err)
	}

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

	_, err = s.slidesService.Presentations.BatchUpdate(presentationID, &slides.BatchUpdatePresentationRequest{
		Requests: requests,
	}).Do()

	if err != nil {
		return fmt.Errorf("error reordering slides: %w", err)
	}

	return nil
}
