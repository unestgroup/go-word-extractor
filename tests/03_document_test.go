package tests

import (
	"testing"
	word_extractor "word-extractor/pkg/word-extractor"

	"github.com/stretchr/testify/assert"
)

func TestDocument(t *testing.T) {
	t.Run("should instantiate successfully", func(t *testing.T) {
		document := word_extractor.NewDocument()
		assert.NotNil(t, document)
	})

	t.Run("should read the body", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Body = "This is the body"
		assert.Equal(t, "This is the body", document.GetBody(nil))
	})

	t.Run("should read the footnotes", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Footnotes = "This is the footnotes"
		assert.Equal(t, "This is the footnotes", document.GetFootnotes(nil))
	})

	t.Run("should read the endnotes", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Endnotes = "This is the endnotes"
		assert.Equal(t, "This is the endnotes", document.GetEndnotes(nil))
	})

	t.Run("should read the annotations", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Annotations = "This is the annotations"
		assert.Equal(t, "This is the annotations", document.GetAnnotations(nil))
	})

	t.Run("should read the headers", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Headers = "This is the headers"
		assert.Equal(t, "This is the headers", document.GetHeaders(nil))
	})

	t.Run("should read the headers and footers", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Headers = "This is the headers\n"
		document.Footers = "This is the footers\n"
		assert.Equal(t, "This is the headers\nThis is the footers\n", document.GetHeaders(nil))
	})

	t.Run("should selectively read the headers", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Headers = "This is the headers\n"
		document.Footers = "This is the footers\n"
		opts := &word_extractor.Options{IncludeFooters: false}
		assert.Equal(t, "This is the headers\n", document.GetHeaders(opts))
	})

	t.Run("should read the footers", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Headers = "This is the headers\n"
		document.Footers = "This is the footers\n"
		assert.Equal(t, "This is the footers\n", document.GetFooters(nil))
	})

	t.Run("should read the body textboxes", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Textboxes = "This is the textboxes\n"
		document.HeaderTextboxes = "This is the header textboxes\n"
		opts := &word_extractor.Options{
			IncludeBody:              true,
			IncludeHeadersAndFooters: false,
		}
		assert.Equal(t, "This is the textboxes\n", document.GetTextboxes(opts))
	})

	t.Run("should read the header textboxes", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Textboxes = "This is the textboxes\n"
		document.HeaderTextboxes = "This is the header textboxes\n"
		opts := &word_extractor.Options{
			IncludeBody:              false,
			IncludeHeadersAndFooters: true,
		}
		assert.Equal(t, "This is the header textboxes\n", document.GetTextboxes(opts))
	})

	t.Run("should read all textboxes", func(t *testing.T) {
		document := word_extractor.NewDocument()
		document.Textboxes = "This is the textboxes\n"
		document.HeaderTextboxes = "This is the header textboxes\n"
		opts := &word_extractor.Options{
			IncludeBody:              true,
			IncludeHeadersAndFooters: true,
		}
		assert.Equal(t, "This is the textboxes\n\nThis is the header textboxes\n", document.GetTextboxes(opts))
	})
}
