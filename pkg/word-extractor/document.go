package word_extractor

import "strings"

// Document implements the main document returned when a Word file has been extracted.
// This exposes methods that allow the body, annotations, headers, footnotes, and
// endnotes to be read and used.
type Document struct {
	Body            string
	Footnotes       string
	Endnotes        string
	Headers         string
	Footers         string
	Annotations     string
	Textboxes       string
	HeaderTextboxes string
}

// Options contains configuration for document content retrieval
type Options struct {
	// FilterUnicode if true (the default), converts common Unicode quotes to ASCII
	FilterUnicode bool
	// IncludeFooters if true (the default), returns headers and footers as a single string
	IncludeFooters bool
	// IncludeHeadersAndFooters if true (the default), includes text box content in headers/footers
	IncludeHeadersAndFooters bool
	// IncludeBody if true (the default), includes text box content in document body
	IncludeBody bool
}

func NewDocument() *Document {
	return &Document{}
}

func defaultOptions() *Options {
	return &Options{
		FilterUnicode:            true,
		IncludeFooters:           true,
		IncludeHeadersAndFooters: true,
		IncludeBody:              true,
	}
}

// GetBody returns the main body part of a Word file
func (d *Document) GetBody(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	if opts.FilterUnicode {
		return Filter(d.Body)
	}
	return d.Body
}

// GetHeaders returns the headers part of a Word file, optionally including footers
func (d *Document) GetHeaders(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	value := d.Headers
	if opts.IncludeFooters {
		value += d.Footers
	}
	if opts.FilterUnicode {
		return Filter(value)
	}
	return value
}

// Other getter methods follow the same pattern
func (d *Document) GetFootnotes(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	if opts.FilterUnicode {
		return Filter(d.Footnotes)
	}
	return d.Footnotes
}

func (d *Document) GetEndnotes(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	if opts.FilterUnicode {
		return Filter(d.Endnotes)
	}
	return d.Endnotes
}

func (d *Document) GetFooters(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	if opts.FilterUnicode {
		return Filter(d.Footers)
	}
	return d.Footers
}

func (d *Document) GetAnnotations(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}
	if opts.FilterUnicode {
		return Filter(d.Annotations)
	}
	return d.Annotations
}

// GetTextboxes returns the textbox content from a Word file
func (d *Document) GetTextboxes(opts *Options) string {
	if opts == nil {
		opts = defaultOptions()
	}

	var parts []string
	if opts.IncludeBody {
		parts = append(parts, d.Textboxes)
	}
	if opts.IncludeHeadersAndFooters {
		parts = append(parts, d.HeaderTextboxes)
	}

	result := strings.Join(parts, "\n")
	if opts.FilterUnicode {
		return Filter(result)
	}
	return result
}
