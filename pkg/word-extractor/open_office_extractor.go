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
	pieces        [][]rune // Changed from []string to [][]rune
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

	case "tab":
		if len(e.context) > 0 && e.context[0] == "content" {
			e.pieces = append(e.pieces, []rune("\t"))
		}

	case "br":
		if len(e.context) > 0 && e.context[0] == "content" {
			e.pieces = append(e.pieces, []rune("\n"))
		}

	case "endnote", "footnote":
		typ := "content"
		for _, attr := range se.Attr {
			if attr.Name.Local == "type" {
				typ = attr.Value
			}
		}
		e.context = append([]string{typ}, e.context...)

	case "del", "instrText":
		e.context = append([]string{"deleted"}, e.context...)

	case "tc":
		e.context = append([]string{"cell"}, e.context...)

	case "drawing":
		e.context = append([]string{"drawing"}, e.context...)

	case "txbxContent":
		e.context = append([]string{"textbox"}, e.context...)
		oldPieces := e.pieces
		e.pieces = [][]rune{}
		e.context = append([]string{"textbox"}, e.context...)
		e.pieces = append(oldPieces, e.pieces...)
	}
}

func (e *OpenOfficeExtractor) handleCloseTag(ee xml.EndElement) {
	// Only check Local name if it's in the Word ML namespace
	if !e.isWordMLElement(ee.Name) && ee.Name.Local != "Override" && ee.Name.Local != "Default" && ee.Name.Local != "Relationship" {
		return
	}

	switch ee.Name.Local {
	case "document":
		e.document.Body = string(joinRunes(e.pieces))
		e.context = nil

	case "footnotes":
		e.document.Footnotes = string(joinRunes(e.pieces))
		e.context = nil

	case "endnotes":
		e.document.Endnotes = string(joinRunes(e.pieces))
		e.context = nil

	case "comments":
		e.document.Annotations = string(joinRunes(e.pieces))
		e.context = nil

	case "hdr":
		e.document.Headers += string(joinRunes(e.pieces))
		e.context = nil

	case "ftr":
		e.document.Footers += string(joinRunes(e.pieces))
		e.context = nil

	case "p":
		if len(e.context) > 0 && (e.context[0] == "content" || e.context[0] == "cell" || e.context[0] == "textbox") {
			e.pieces = append(e.pieces, []rune("\n"))
		}

	case "tc":
		if len(e.pieces) > 0 {
			e.pieces = e.pieces[:len(e.pieces)-1]
		}
		e.pieces = append(e.pieces, []rune("\t"))
		if len(e.context) > 0 {
			e.context = e.context[1:]
		}

	case "txbxContent":
		if len(e.context) < 3 {
			return
		}
		textBox := string(joinRunes(e.pieces))
		e.context = e.context[1:] // remove "textbox"

		// Restore previous pieces from context
		prevPieces := [][]rune{}
		idx := 0
		for i, ctx := range e.context {
			if ctx != "" {
				prevPieces = e.pieces[i:]
				idx = i
				break
			}
		}
		e.context = e.context[:idx]
		e.pieces = prevPieces

		// If in drawing context, discard
		if len(e.context) > 0 && e.context[0] == "drawing" {
			return
		}

		if textBox == "" {
			return
		}

		inHeader := false
		for _, ctx := range e.context {
			if ctx == "header" {
				inHeader = true
				break
			}
		}

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

	if e.context[0] == "content" || e.context[0] == "cell" || e.context[0] == "textbox" {
		e.pieces = append(e.pieces, []rune(string(cd)))
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
