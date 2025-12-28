package presentation

import (
	"context"
	"fmt"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/slides/v1"
)

// Service wraps Google Slides and Drive services for presentation operations.
type Service struct {
	slidesService *slides.Service
	driveService  *drive.Service
}

// NewService creates a new presentation service.
func NewService(ctx context.Context, slidesService *slides.Service, driveService *drive.Service) *Service {
	return &Service{
		slidesService: slidesService,
		driveService:  driveService,
	}
}

// Create creates a new presentation with the given title.
func (s *Service) Create(ctx context.Context, title string, folderID string) (*slides.Presentation, error) {
	presentation := &slides.Presentation{
		Title: title,
	}

	result, err := s.slidesService.Presentations.Create(presentation).Do()
	if err != nil {
		return nil, fmt.Errorf("error creating presentation: %w", err)
	}

	if folderID != "" {
		_, err = s.driveService.Files.Update(result.PresentationId, &drive.File{}).
			AddParents(folderID).Do()
		if err != nil {
			return nil, fmt.Errorf("error moving to folder: %w", err)
		}
	}

	return result, nil
}

// Get retrieves a presentation by ID.
func (s *Service) Get(ctx context.Context, presentationID string) (*slides.Presentation, error) {
	presentation, err := s.slidesService.Presentations.Get(presentationID).Do()
	if err != nil {
		return nil, fmt.Errorf("error getting presentation: %w", err)
	}
	return presentation, nil
}
