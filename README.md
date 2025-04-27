# go-word-extractor

[![Go Report Card](https://goreportcard.com/badge/github.com/unestgroup/go-word-extractor)](https://goreportcard.com/report/github.com/unestgroup/go-word-extractor)
[![Go Reference](https://pkg.go.dev/badge/github.com/unestgroup/go-word-extractor.svg)](https://pkg.go.dev/github.com/unestgroup/go-word-extractor)

Read text content from Word documents (.doc and .docx) using Go. Ported from the Node.js [word-extractor](https://github.com/morungos/node-word-extractor) library.

## Why use this module?

There are various ways to extract text from Word files, but they often require external programs or libraries, increasing installation and runtime complexity.

This module provides a fast way to read text from Word files entirely within the Go environment.

*   **No External Dependencies:** You don't need Word, Office, or any other external software installed.
*   **Cross-Platform:** Works on any platform supported by Go.
*   **Pure Go:** No CGo or native binary requirements.
*   **Supports .doc and .docx:** Handles both traditional OLE-based (.doc) and modern Open Office XML (.docx) formats.
*   **Flexible Input:** Works with file paths or `[]byte` slices.

## How do I install this module?

```bash
# Replace 'your-username/go-word-extractor' with the actual repository path
go get github.com/unestgroup/go-word-extractor
```

## How do I use this module?

### As a Library

```go
package main

import (
	"fmt"
	"log"

	// Replace with the actual import path
	word_extractor "github.com/unestgroup/go-word-extractor/pkg/word-extractor"
)

func main() {
	extractor := word_extractor.NewWordExtractor()
	filePath := "path/to/your/file.docx" // Or file.doc

	doc, err := extractor.Extract(filePath)
	if err != nil {
		log.Fatalf("Error extracting content: %v", err)
	}

	// Get different parts of the document
	body := doc.GetBody(nil) // Pass nil for default options
	fmt.Println("--- Body ---")
	fmt.Println(body)

	headers := doc.GetHeaders(nil) // Pass nil for default options
	fmt.Println("\n--- Headers ---")
	fmt.Println(headers)

	// You can also extract footnotes, endnotes, annotations, etc.
	// footnotes := doc.GetFootnotes(nil)
	// endnotes := doc.GetEndnotes(nil)
	// annotations := doc.GetAnnotations(nil)
	// textboxes := doc.GetTextboxes(nil) // Note: Options might differ from Node version
}

```

The `Extract` method returns a `*Document` object (or an error), which provides methods to access different parts of the document.

### As a Command-Line Tool

You can build and run the tool from the command line:

1.  **Build:**
    ```bash
    go build -o word-extractor-cli .
    ```
2.  **Run:**
    ```bash
    # Process a single file
    ./word-extractor-cli path/to/your/document.docx

    # Process multiple files/directories
    ./word-extractor-cli file1.doc folder1 file2.docx

    # Process directories recursively
    ./word-extractor-cli -r folder1 folder2
    ```
    The tool will print the extracted headers and body content for each processed file to the standard output. Errors encountered during processing will be logged.

## API

### `word_extractor.NewWordExtractor() *WordExtractor`

Creates a new instance of the extractor.

### `WordExtractor.Extract(source interface{}) (*Document, error)`

Main method to open and process a Word file.
*   `source`: Can be either a file path (`string`) or the file content as a `[]byte` slice.
*   Returns a `*Document` pointer on success, or an `error` if extraction fails.

### `Document.GetBody(options map[string]interface{}) string`

Retrieves the main content text from the document. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed).

### `Document.GetHeaders(options map[string]interface{}) string`

Retrieves header and footer text. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed). *Note: Unlike the Node.js version, this currently retrieves both headers and footers combined. Specific options for separation might be added later.*

### `Document.GetFootnotes(options map[string]interface{}) string`

Retrieves footnote text. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed).

### `Document.GetEndnotes(options map[string]interface{}) string`

Retrieves endnote text. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed).

### `Document.GetAnnotations(options map[string]interface{}) string`

Retrieves comment bubble text. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed).

### `Document.GetTextboxes(options map[string]interface{}) string`

Retrieves textbox content. Handles UNICODE characters correctly.
*   `options`: A map for potential future options (currently `nil` can be passed). *Note: Options for including/excluding body or header/footer textboxes might differ from the Node.js version.*

## License

Licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
