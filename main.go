package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	word_extractor "word-extractor/pkg/word-extractor"
)

func main() {
	// Create a new extractor
	extractor := word_extractor.NewWordExtractor()

	// Test with a file path
	filePath := "../__tests__/data/test06.docx" // Make sure to have a test Word document in this path
	// Validate file path
	absPath, err := filepath.Abs(filePath)
	if err != nil {
		log.Fatalf("Error resolving file path: %v", err)
	}

	// Check if file exists
	if _, err := os.Stat(absPath); os.IsNotExist(err) {
		log.Fatalf("File does not exist at path: %s", absPath)
	}

	doc, err := extractor.Extract(filePath)
	if err != nil {
		log.Fatalf("Error extracting content: %v", err)
	}

	// Print the extracted content
	fmt.Println("Extracted content:")

	// for _, r := range doc.GetBody(nil) {
	// 	// fmt.Printf("%v %c\n", r, r)
	// }
	fmt.Println(doc.GetBody(nil))

	// Test with byte slice (if you have the file content in memory)
	// fileBytes, err := os.ReadFile(filePath)
	// if err != nil {
	//     log.Fatalf("Error reading file: %v", err)
	// }
	//
	// doc, err = extractor.Extract(fileBytes)
	// if err != nil {
	//     log.Fatalf("Error extracting content: %v", err)
	// }
}
