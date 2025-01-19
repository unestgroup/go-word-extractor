package word_extractor

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
)

// WordExtractor is the main struct for extracting content from Word documents
type WordExtractor struct{}

// NewWordExtractor creates a new instance of WordExtractor
func NewWordExtractor() *WordExtractor {
	return &WordExtractor{}
}

// Extract processes the given source (either filename or byte slice) and extracts the document content
func (w *WordExtractor) Extract(source interface{}) (*Document, error) {
	var reader io.ReadSeeker
	var closer io.Closer

	switch s := source.(type) {
	case []byte:
		reader = bytes.NewReader(s)
	case string:
		// Get absolute path
		absPath, err := filepath.Abs(s)
		if err != nil {
			return nil, fmt.Errorf("failed to get absolute path: %v", err)
		}

		// Open file with explicit read permissions
		file, err := os.OpenFile(absPath, os.O_RDONLY, 0644)
		if err != nil {
			return nil, fmt.Errorf("failed to open file: %v", err)
		}
		reader = file
		closer = file
	default:
		return nil, errors.New("source must be either a filename string or byte slice")
	}

	if closer != nil {
		defer closer.Close()
	}

	// Read first 512 bytes to determine file type
	buffer := make([]byte, 512)
	_, err := reader.Read(buffer)
	if err != nil {
		return nil, err
	}

	// Reset reader position
	_, err = reader.Seek(0, 0)
	if err != nil {
		return nil, err
	}

	// Check file signature
	if len(buffer) < 4 {
		return nil, errors.New("file too small")
	}

	var extractor DocumentExtractor

	// Check for OLE document (0xD0CF)
	if binary.BigEndian.Uint16(buffer[0:2]) == 0xD0CF {
		extractor = NewWordOleExtractor()
	} else if binary.BigEndian.Uint16(buffer[0:2]) == 0x504B {
		// Check for OpenOffice document (PK signature + specific following bytes)
		next := binary.BigEndian.Uint16(buffer[2:4])
		if next == 0x0304 || next == 0x0506 || next == 0x0708 {
			extractor = NewOpenOfficeExtractor()
		}
	}

	if extractor == nil {
		return nil, errors.New("unable to read this type of file")
	}

	return extractor.Extract(reader)
}

// DocumentExtractor interface defines the contract for different document extractors
type DocumentExtractor interface {
	Extract(reader io.ReadSeeker) (*Document, error)
}
