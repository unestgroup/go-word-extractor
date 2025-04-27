package tests

import (
	"io/ioutil"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
	word_extractor "word-extractor/pkg/word-extractor"

	"github.com/stretchr/testify/require"
)

// cleanHeaderText removes extra whitespace and normalizes newlines
func cleanHeaderText(text string) string {
	text = strings.TrimSpace(text)
	multiNewlines := regexp.MustCompile(`\n{2,}`)
	return multiNewlines.ReplaceAllString(text, "\n\n")
}

func TestOpenOfficeFilesExtract(t *testing.T) {
	// Get test data directory
	dataDir := filepath.Join(".", "data")
	files, err := ioutil.ReadDir(dataDir)
	require.NoError(t, err)

	// Find matching .doc and .docx pairs
	docxFiles := make(map[string]bool)
	var testFiles []string

	for _, f := range files {
		if strings.HasSuffix(f.Name(), ".docx") {
			docxFiles[f.Name()] = true
		}
	}

	for _, f := range files {
		if strings.HasPrefix(f.Name(), "~") {
			continue
		}
		if !strings.HasSuffix(f.Name(), ".doc") {
			continue
		}
		if docxFiles[f.Name()+"x"] {
			testFiles = append(testFiles, f.Name())
		}
	}

	extractor := word_extractor.NewWordExtractor()

	for _, file := range testFiles {
		t.Run(file, func(t *testing.T) {
			docPath := filepath.Join(dataDir, file)
			docxPath := filepath.Join(dataDir, file+"x")

			// Extract content from both files
			docDoc, err := extractor.Extract(docPath)
			require.NoError(t, err)

			docxDoc, err := extractor.Extract(docxPath)
			require.NoError(t, err)

			// Compare content between .doc and .docx versions
			opts := &word_extractor.Options{FilterUnicode: true}

			// Compare main body content
			require.Equal(t,
				normalizeText(docDoc.GetBody(opts)),
				normalizeText(docxDoc.GetBody(opts)),
				"Body content mismatch",
			)

			// Compare footnotes
			require.Equal(t,
				normalizeText(docDoc.GetFootnotes(opts)),
				normalizeText(docxDoc.GetFootnotes(opts)),
				"Footnotes mismatch",
			)

			// Compare endnotes
			require.Equal(t,
				normalizeText(docDoc.GetEndnotes(opts)),
				normalizeText(docxDoc.GetEndnotes(opts)),
				"Endnotes mismatch",
			)

			// Compare annotations
			require.Equal(t,
				normalizeText(docDoc.GetAnnotations(opts)),
				normalizeText(docxDoc.GetAnnotations(opts)),
				"Annotations mismatch",
			)

			// Compare textboxes (excluding headers/footers)
			require.Equal(t,
				normalizeText(docDoc.GetTextboxes(&word_extractor.Options{
					FilterUnicode:            true,
					IncludeHeadersAndFooters: false,
				})),
				normalizeText(docxDoc.GetTextboxes(&word_extractor.Options{
					FilterUnicode:            true,
					IncludeHeadersAndFooters: false,
				})),
				"Textboxes mismatch",
			)

			// Compare header textboxes
			require.Equal(t,
				normalizeText(docDoc.GetTextboxes(&word_extractor.Options{
					FilterUnicode: true,
					IncludeBody:   false,
				})),
				normalizeText(docxDoc.GetTextboxes(&word_extractor.Options{
					FilterUnicode: true,
					IncludeBody:   false,
				})),
				"Header textboxes mismatch",
			)
		})
	}
}

// normalizeText standardizes text for comparison
func normalizeText(text string) string {
	// Remove multiple spaces
	spaceRegex := regexp.MustCompile(`\s+`)
	text = spaceRegex.ReplaceAllString(text, " ")

	// Remove multiple newlines
	newlineRegex := regexp.MustCompile(`\n+`)
	text = newlineRegex.ReplaceAllString(text, "\n")

	// Trim spaces and normalize line endings
	text = strings.TrimSpace(text)
	text = strings.ReplaceAll(text, "\r\n", "\n")

	return text
}
