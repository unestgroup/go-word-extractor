package word_extractor

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strings"
)

type OpenOfficeExtractor struct {
	document      *Document
	streamTypes   map[string]bool
	headerTypes   map[string]bool
	actions       map[string]Action
	defaults      map[string]string
	relationships map[string]Relationship
	context       []string
	pieces        [][]rune   // Changed from []string to [][]rune
	piecesStack   [][][]rune // Stack to hold pieces state for nested contexts like textboxes
}

type Action struct {
	action interface{}
	typ    string
}

type Relationship struct {
	Type   string
	Target string
}

func NewOpenOfficeExtractor() *OpenOfficeExtractor {
	e := &OpenOfficeExtractor{
		streamTypes: map[string]bool{
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":    true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml":         true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml": true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":        true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":         true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml":           true,
			"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml":           true,
			"application/vnd.openxmlformats-package.relationships+xml":                            true,
		},
		headerTypes: map[string]bool{
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header": true,
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer": true,
		},
		actions:       make(map[string]Action),
		defaults:      make(map[string]string),
		relationships: make(map[string]Relationship),
	}
	return e
}

// Update the Extract method signature to match the interface
func (e *OpenOfficeExtractor) Extract(reader io.ReadSeeker) (*Document, error) {
	e.document = NewDocument()
	e.relationships = make(map[string]Relationship)

	size, err := reader.Seek(0, io.SeekEnd)
	if err != nil {
		return nil, err
	}

	// Reset to beginning of file
	_, err = reader.Seek(0, io.SeekStart)
	if err != nil {
		return nil, err
	}

	zr, err := zip.NewReader(reader.(io.ReaderAt), size)
	if err != nil {
		return nil, err
	}

	// Build entry table and order files
	entryTable := make(map[string]*zip.File)
	entryNames := make([]string, 0)

	for _, f := range zr.File {
		entryTable[f.Name] = f
		entryNames = append(entryNames, f.Name)
	}

	// Process [Content_Types].xml first
	contentTypesFile := "[Content_Types].xml"
	found := false
	for i, name := range entryNames {
		if name == contentTypesFile {
			found = true
			// Remove and prepend to front
			entryNames = append(entryNames[:i], entryNames[i+1:]...)
			entryNames = append([]string{contentTypesFile}, entryNames...)
			e.actions[contentTypesFile] = Action{typ: "content-types"}
			break
		}
	}
	if !found {
		return nil, fmt.Errorf("invalid Open Office XML: missing content types")
	}

	// Process entries in order
	for _, name := range entryNames {
		if e.shouldProcess(name) {
			if err := e.handleEntry(entryTable[name]); err != nil {
				return nil, err
			}
		}
	}

	// Post-process textboxes and headerTextboxes
	if e.document.Textboxes != "" {
		e.document.Textboxes += "\n"
	}
	if e.document.HeaderTextboxes != "" {
		e.document.HeaderTextboxes += "\n"
	}

	return e.document, nil
}

func (e *OpenOfficeExtractor) shouldProcess(filename string) bool {
	if _, ok := e.actions[filename]; ok {
		return true
	}
	ext := path.Ext(filename)
	if ext == "" {
		return false
	}
	defaultType, ok := e.defaults[ext[1:]]
	return ok && e.streamTypes[defaultType]
}

func (e *OpenOfficeExtractor) handleEntry(f *zip.File) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		switch t := token.(type) {
		case xml.StartElement:
			e.handleOpenTag(t)
		case xml.EndElement:
			e.handleCloseTag(t)
		case xml.CharData:
			e.handleCharData(t)
		}
	}
	return nil
}

const (
	WordMLNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
)

func (e *OpenOfficeExtractor) isWordMLElement(se xml.Name) bool {
	return se.Space == WordMLNamespace
}

func (e *OpenOfficeExtractor) handleOpenTag(se xml.StartElement) {
	// For debugging
	// fmt.Printf("StartElement Space: %s, Local: %s\n", se.Name.Space, se.Name.Local)

	// Only check Local name if it's in the Word ML namespace
	if !e.isWordMLElement(se.Name) && se.Name.Local != "Override" && se.Name.Local != "Default" && se.Name.Local != "Relationship" {
		return
	}

	switch se.Name.Local {
	// Match JS order
	case "Override":
		var contentType, partName string
		for _, attr := range se.Attr {
			switch attr.Name.Local {
			case "ContentType":
				contentType = attr.Value
			case "PartName":
				partName = strings.TrimPrefix(attr.Value, "/")
			}
		}
		if e.streamTypes[contentType] {
			e.actions[partName] = Action{typ: contentType, action: e.streamTypes[contentType]}
		}

	case "Default":
		var extension, contentType string
		for _, attr := range se.Attr {
			switch attr.Name.Local {
			case "Extension":
				extension = attr.Value
			case "ContentType":
				contentType = attr.Value
			}
		}
		e.defaults[extension] = contentType

	case "Relationship":
		var id, typ, target string
		for _, attr := range se.Attr {
			switch attr.Name.Local {
			case "Id":
				id = attr.Value
			case "Type":
				typ = attr.Value
			case "Target":
				target = attr.Value
			}
		}
		e.relationships[id] = Relationship{Type: typ, Target: target}

	case "document", "footnotes", "endnotes", "comments":
		e.context = []string{"content", "body"}
		e.pieces = [][]rune{}

	case "hdr", "ftr":
		e.context = []string{"content", "header"}
		e.pieces = [][]rune{}

	case "endnote", "footnote": // JS: w:endnote, w:footnote
		typ := "content"
		for _, attr := range se.Attr {
			if attr.Name.Local == "type" {
				typ = attr.Value
			}
		}
		e.context = append([]string{typ}, e.context...)

	case "tab": // JS: w:tab
		if len(e.context) > 0 && e.context[0] == "content" {
			e.pieces = append(e.pieces, []rune("\t"))
		}

	case "br": // JS: w:br
		if len(e.context) > 0 && e.context[0] == "content" {
			// Check for page break type, although currently treated the same as line break
			// brType := ""
			// for _, attr := range se.Attr {
			// 	if attr.Name.Local == "type" {
			// 		brType = attr.Value
			// 		break
			// 	}
			// }
			// if brType == "page" {
			// 	e.pieces = append(e.pieces, []rune("\\n")) // Or handle page break differently if needed later
			// } else {
			// 	e.pieces = append(e.pieces, []rune("\\n"))
			// }
			// Simplified version since outcome is currently the same:
			e.pieces = append(e.pieces, []rune("\n"))
		}

	case "del", "instrText": // JS: w:del, w:instrText
		e.context = append([]string{"deleted"}, e.context...)

	case "tabs": // JS: w:tabs
		e.context = append([]string{"tabs"}, e.context...)

	case "tc": // JS: w:tc
		e.context = append([]string{"cell"}, e.context...)

	case "drawing": // JS: w:drawing
		e.context = append([]string{"drawing"}, e.context...)

	case "txbxContent": // JS: w:txbxContent
		// Push current pieces onto the stack
		e.piecesStack = append(e.piecesStack, e.pieces)
		// Reset pieces for the textbox content
		e.pieces = [][]rune{}
		// Push textbox context marker
		e.context = append([]string{"textbox"}, e.context...)
		// --- Original incorrect Go logic removed ---
		// oldPieces := e.pieces
		// e.pieces = [][]rune{}
		// e.context = append([]string{"textbox"}, e.context...)
		// e.pieces = append(oldPieces, e.pieces...)
	}
}

func (e *OpenOfficeExtractor) handleCloseTag(ee xml.EndElement) {
	// Only check Local name if it's in the Word ML namespace
	if !e.isWordMLElement(ee.Name) && ee.Name.Local != "Override" && ee.Name.Local != "Default" && ee.Name.Local != "Relationship" {
		return
	}

	switch ee.Name.Local {
	// Match JS order
	case "document": // JS: w:document
		e.document.Body = string(joinRunes(e.pieces))
		e.context = nil

	case "footnote", "endnote": // JS: w:footnote, w:endnote (Combined in Go)
		if len(e.context) > 0 {
			e.context = e.context[1:]
		}

	case "footnotes": // JS: w:footnotes
		e.document.Footnotes = string(joinRunes(e.pieces))
		e.context = nil

	case "endnotes": // JS: w:endnotes
		e.document.Endnotes = string(joinRunes(e.pieces))
		e.context = nil

	case "comments": // JS: w:comments
		e.document.Annotations = string(joinRunes(e.pieces))
		e.context = nil

	case "hdr": // JS: w:hdr
		e.document.Headers += string(joinRunes(e.pieces))
		e.context = nil

	case "ftr": // JS: w:ftr
		e.document.Footers += string(joinRunes(e.pieces))
		e.context = nil

	case "p": // JS: w:p
		if len(e.context) > 0 && (e.context[0] == "content" || e.context[0] == "cell" || e.context[0] == "textbox") {
			e.pieces = append(e.pieces, []rune("\n")) // Corrected: Use newline rune
		}

	case "del", "instrText": // JS: w:del, w:instrText
		if len(e.context) > 0 {
			e.context = e.context[1:]
		}

	case "tabs": // JS: w:tabs
		if len(e.context) > 0 {
			e.context = e.context[1:]
		}

	case "tc":
		// In JS, it pops the last piece (often a \n from <w:p>) before adding \t.
		if len(e.pieces) > 0 {
			e.pieces = e.pieces[:len(e.pieces)-1]
		}
		e.pieces = append(e.pieces, []rune("\t"))
		if len(e.context) > 0 {
			e.context = e.context[1:] // Pop "cell"
		}

	case "tr": // JS: w:tr
		// Add newline after a table row (Matches JS unconditional behavior)
		e.pieces = append(e.pieces, []rune("\n")) // Corrected: Use newline rune and remove condition

	case "drawing": // JS: w:drawing
		if len(e.context) > 0 {
			e.context = e.context[1:]
		}

	case "txbxContent":
		// Get the text content accumulated within the textbox
		textBox := string(joinRunes(e.pieces))

		if e.context[0] != "textbox" {
			fmt.Printf("Warning: Invalid textbox context\n")
			return
		}
		// Pop the textbox context marker
		if len(e.context) > 0 {
			e.context = e.context[1:] // Pop "textbox"
		}

		// Pop the previous pieces state from the stack
		if len(e.piecesStack) > 0 {
			e.pieces = e.piecesStack[len(e.piecesStack)-1]
			e.piecesStack = e.piecesStack[:len(e.piecesStack)-1]
		} else {
			// Should not happen if open/close tags are balanced
			e.pieces = [][]rune{} // Reset pieces if stack is empty
		}

		// --- Original incorrect Go restoration logic removed ---
		// if len(e.context) < 3 { // This check was arbitrary
		// 	return
		// }
		// textBox := string(joinRunes(e.pieces))
		// e.context = e.context[1:] // remove "textbox"
		// // Restore previous pieces from context
		// prevPieces := [][]rune{}
		// idx := 0
		// for i, ctx := range e.context {
		// 	if ctx != "" {
		// 		prevPieces = e.pieces[i:] // Incorrect logic
		// 		idx = i
		// 		break
		// 	}
		// }
		// e.context = e.context[:idx]
		// e.pieces = prevPieces

		// If in drawing context, discard (Matches JS)
		if len(e.context) > 0 && e.context[0] == "drawing" {
			return
		}

		if textBox == "" { // Matches JS
			return
		}

		// Check if inside a header/footer (Matches JS)
		inHeader := false
		for _, ctx := range e.context {
			if ctx == "header" {
				inHeader = true
				break
			}
		}

		// Append to the appropriate document field (Matches JS)
		if inHeader {
			if e.document.HeaderTextboxes != "" {
				e.document.HeaderTextboxes += "\n"
			}
			e.document.HeaderTextboxes += textBox
		} else {
			if e.document.Textboxes != "" {
				e.document.Textboxes += "\n"
			}
			e.document.Textboxes += textBox
		}
	}
}

func (e *OpenOfficeExtractor) handleCharData(cd xml.CharData) {
	if len(e.context) == 0 {
		return
	}

	// fmt.Printf("CharData: %s\n", string(cd))
	// fmt.Printf("Current context: %s\n", e.context[0])

	if e.context[0] == "content" || e.context[0] == "cell" || e.context[0] == "textbox" {
		// fmt.Printf("Append to pieces: %s\n", string(cd))
		e.pieces = append(e.pieces, []rune(string(cd)))

		// fmt.Printf("Current pieces: %s\n", string(joinRunes(e.pieces)))
	}
}

// Helper function to join rune slices
func joinRunes(pieces [][]rune) []rune {
	var total int
	for _, p := range pieces {
		total += len(p)
	}

	result := make([]rune, 0, total)
	for _, p := range pieces {
		result = append(result, p...)
	}
	return result
}
