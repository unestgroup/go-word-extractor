package word_extractor

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"regexp"
	"strings"
	"unicode/utf16"
)

// Binary to Unicode conversion table
var binaryToUnicodeTable = map[rune]string{
	0x0082: "\u201a", // single low-9 quotation mark
	0x0083: "\u0192", // latin small letter f with hook
	0x0084: "\u201e", // double low-9 quotation mark
	0x0085: "\u2026", // horizontal ellipsis
	0x0086: "\u2020", // dagger
	0x0087: "\u2021", // double dagger
	0x0088: "\u02C6", // modifier letter circumflex accent
	0x0089: "\u2030", // per mille sign
	0x008a: "\u0160", // latin capital letter s with caron
	0x008b: "\u2039", // single left-pointing angle quotation mark
	0x008c: "\u0152", // latin capital ligature oe
	0x008e: "\u017D", // latin capital letter z with caron
	0x0091: "\u2018", // left single quotation mark
	0x0092: "\u2019", // right single quotation mark
	0x0093: "\u201C", // left double quotation mark
	0x0094: "\u201D", // right double quotation mark
	0x0095: "\u2022", // bullet
	0x0096: "\u2013", // en dash
	0x0097: "\u2014", // em dash
	0x0098: "\u02DC", // small tilde
	0x0099: "\u2122", // trade mark sign
	0x009a: "\u0161", // latin small letter s with caron
	0x009b: "\u203A", // single right-pointing angle quotation mark
	0x009c: "\u0153", // latin small ligature oe
	0x009e: "\u017E", // latin small letter z with caron
	0x009f: "\u0178", // latin capital letter y with diaeresis
}

// Character replacement table for cleaning
var replaceTable = map[rune]string{
	0x0002: "\x00",
	0x0005: "\x00",
	0x0007: "\t",
	0x0008: "\x00",
	0x000A: "\n",
	0x000B: "\n",
	0x000C: "\n",
	0x000D: "\n",
	0x001E: "\u2011",
}

// Filter table for standard punctuation
var filterTable = map[rune]string{
	0x2002: " ",  // en space
	0x2003: " ",  // em space
	0x2012: "-",  // figure dash
	0x2013: "-",  // en dash
	0x2014: "-",  // em dash
	0x2018: "'",  // left single quote
	0x2019: "'",  // right single quote
	0x201c: "\"", // left double quote
	0x201d: "\"", // right double quote
}

var fieldRegex = regexp.MustCompile(`(?:\x13[^\x13\x14\x15]*\x14?([^\x13\x14\x15]*)\x15)`)

func bufferToUCS2String(buffer []byte) (string, error) {
	// Ensure the buffer length is even, as UCS-2 characters are 2 bytes each
	if len(buffer)%2 != 0 {
		return "", fmt.Errorf("buffer length must be even, got %d", len(buffer))
	}

	// Read uint16 values from the buffer
	uint16Buffer := make([]uint16, len(buffer)/2)
	err := binary.Read(bytes.NewReader(buffer), binary.LittleEndian, &uint16Buffer)
	if err != nil {
		return "", err
	}
	// Decode the uint16 array into a UTF-16 string
	return string(utf16.Decode(uint16Buffer)), nil
}

// binaryToUnicode converts []byte Windows-1252 encoded text to Unicode
func binaryToUnicode(text string) string {
	var result strings.Builder

	for _, r := range text {
		if replacement, ok := binaryToUnicodeTable[r]; ok {
			result.WriteString(replacement)
		} else {
			result.WriteString(string(r))
		}
	}
	return result.String()
}

// Filter replaces common Unicode punctuation with ASCII equivalents
func filterText(text string) string {
	var result strings.Builder
	for _, r := range text {
		if replacement, ok := filterTable[r]; ok {
			result.WriteString(replacement)
		} else {
			result.WriteRune(r)
		}
	}
	return result.String()
}

// CleanText cleans Word document text by handling special characters and fields
func cleanText(text string) string {
	replaceRegex := regexp.MustCompile(`[\x02\x05\x07\x08\x0a\x0b\x0c\x0d\x1f]`)
	text = replaceRegex.ReplaceAllStringFunc(text, func(match string) string {
		if match[0] > 0 {
			if replacement, ok := replaceTable[rune(match[0])]; ok {
				return replacement
			} else {
				return ""
			}
		}
		return match
	})

	called := true
	for called {
		called = false
		text = fieldRegex.ReplaceAllStringFunc(text, func(match string) string {
			called = true
			groups := fieldRegex.FindStringSubmatch(match)
			if len(groups) > 1 {
				return groups[1]
			}
			return match
		})
	}

	// Remove remaining control characters
	text = strings.Map(func(r rune) rune {
		if r <= 0x07 {
			return -1
		}
		return r
	}, text)

	return text
}

// Helper function to check if text contains non-whitespace characters
func containsNonWhitespace(s string) bool {
	for _, r := range s {
		if r != '\r' && r != '\n' && !(r >= '\x02' && r <= '\x08') {
			return true
		}
	}
	return false
}

// join concatenates strings with a separator
func join(strs []string, sep string) string {
	return strings.Join(strs, sep)
}
