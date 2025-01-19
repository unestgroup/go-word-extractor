package tests

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
	word_extractor "word-extractor/pkg/word-extractor"
)

func TestCheckingBlockFromFiles(t *testing.T) {
	extractor := word_extractor.NewWordExtractor()

	t.Run("should extract a .doc document successfully", func(t *testing.T) {
		docPath := filepath.Join("data", "test01.doc")
		_, err := extractor.Extract(docPath)
		if err != nil {
			t.Errorf("Failed to extract .doc file: %v", err)
		}
	})

	t.Run("should extract a .docx document successfully", func(t *testing.T) {
		docxPath := filepath.Join("data", "test01.docx")
		_, err := extractor.Extract(docxPath)
		if err != nil {
			t.Errorf("Failed to extract .docx file: %v", err)
		}
	})

	t.Run("should handle missing file error correctly", func(t *testing.T) {
		missingPath := filepath.Join("data", "missing00.docx")
		_, err := extractor.Extract(missingPath)
		if err == nil {
			t.Error("Expected an error for missing file, but got none")
		}
		expectedErrMsg := "failed to open file"
		if err != nil && !strings.Contains(err.Error(), expectedErrMsg) {
			t.Errorf("Expected error containing '%s', but got: %v", expectedErrMsg, err)
		}
	})

	t.Run("should properly handle file operations", func(t *testing.T) {
		docPath := filepath.Join("data", "test01.doc")

		// Open the file first to verify it exists
		file, err := os.Open(docPath)
		if err != nil {
			t.Skip("Test file not available:", docPath)
		}
		file.Close()

		// Now test the extraction
		_, err = extractor.Extract(docPath)
		if err != nil {
			t.Errorf("Failed to extract document: %v", err)
		}

		// Try to open the file again to verify it was properly closed
		file, err = os.Open(docPath)
		if err != nil {
			t.Errorf("Failed to re-open file after extraction: %v", err)
		}
		file.Close()
	})
}
