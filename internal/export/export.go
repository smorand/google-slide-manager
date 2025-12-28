package export

import (
	"context"
	"fmt"
	"os"

	"google.golang.org/api/drive/v3"
)

// Service wraps Google Drive service for export operations.
type Service struct {
	driveService *drive.Service
}

// NewService creates a new export service.
func NewService(ctx context.Context, driveService *drive.Service) *Service {
	return &Service{
		driveService: driveService,
	}
}

// ToPDF exports a presentation as PDF.
func (s *Service) ToPDF(ctx context.Context, presentationID string, outputFile string) error {
	resp, err := s.driveService.Files.Export(presentationID, "application/pdf").Download()
	if err != nil {
		return fmt.Errorf("error exporting as PDF: %w", err)
	}
	defer resp.Body.Close()

	f, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("error creating output file: %w", err)
	}
	defer f.Close()

	_, err = f.ReadFrom(resp.Body)
	if err != nil {
		return fmt.Errorf("error writing PDF: %w", err)
	}

	return nil
}

// ToPPTX exports a presentation as PowerPoint.
func (s *Service) ToPPTX(ctx context.Context, presentationID string, outputFile string) error {
	resp, err := s.driveService.Files.Export(presentationID, "application/vnd.openxmlformats-officedocument.presentationml.presentation").Download()
	if err != nil {
		return fmt.Errorf("error exporting as PPTX: %w", err)
	}
	defer resp.Body.Close()

	f, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("error creating output file: %w", err)
	}
	defer f.Close()

	_, err = f.ReadFrom(resp.Body)
	if err != nil {
		return fmt.Errorf("error writing PPTX: %w", err)
	}

	return nil
}
