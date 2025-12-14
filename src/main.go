package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "google-slide-manager",
	Short: "Google Slides Manager",
	Long:  "Comprehensive Google Slides operations: create, edit, format, translate, and export presentations",
}

func main() {
	// Presentation operations
	createPresentationCmd.Flags().StringVar(&createPresentationFolderID, "folder", "", "Folder ID to create presentation in")
	rootCmd.AddCommand(createPresentationCmd)

	// Slide operations
	addSlideCmd.Flags().StringVar(&addSlideLayout, "layout", "BLANK", "Slide layout (BLANK, TITLE, TITLE_AND_BODY, etc.)")
	addSlideCmd.Flags().IntVar(&addSlidePosition, "position", -1, "Position to insert slide (-1 for end)")
	rootCmd.AddCommand(addSlideCmd)
	rootCmd.AddCommand(duplicateSlideCmd)
	rootCmd.AddCommand(removeSlideCmd)
	rootCmd.AddCommand(reorderSlidesCmd)
	rootCmd.AddCommand(moveSlideCmd)

	// Table operations
	rootCmd.AddCommand(createTableCmd)
	rootCmd.AddCommand(updateCellCmd)
	styleCellCmd.Flags().StringVar(&styleCellBgColor, "bg-color", "", "Background color (hex, e.g., #FF0000)")
	rootCmd.AddCommand(styleCellCmd)

	// Style operations
	rootCmd.AddCommand(copyTextStyleCmd)
	rootCmd.AddCommand(copyThemeCmd)

	// Notes operations
	rootCmd.AddCommand(getNotesCmd)
	rootCmd.AddCommand(addNotesCmd)
	rootCmd.AddCommand(extractAllNotesCmd)

	// Text operations
	rootCmd.AddCommand(replaceTextCmd)
	rootCmd.AddCommand(extractAllTextCmd)
	rootCmd.AddCommand(searchTextCmd)

	// Shape operations
	rootCmd.AddCommand(addShapeCmd)

	// Translation
	rootCmd.AddCommand(translateSlidesCmd)

	// Export
	rootCmd.AddCommand(exportPdfCmd)
	rootCmd.AddCommand(exportPptxCmd)

	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
