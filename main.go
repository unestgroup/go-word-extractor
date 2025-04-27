package main

import (
	"errors" // Import errors package
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync" // Import the sync package

	word_extractor "word-extractor/pkg/word-extractor"
)

func main() {
	recursive := flag.Bool("r", false, "Recursively search for files in directories")
	flag.Parse()

	// Call the run function with parsed arguments and flags
	if err := run(flag.Args(), *recursive); err != nil {
		log.Fatalf("Error: %v", err) // Log fatal error and exit
	}
}

// run encapsulates the main application logic
func run(args []string, recursive bool) error {
	if len(args) == 0 {
		// Return an error instead of printing usage and exiting directly
		return errors.New("usage: go-word-extractor [-r] <file_or_dir1> [file_or_dir2]")
	}

	extractor := word_extractor.NewWordExtractor()
	var wg sync.WaitGroup // Create a WaitGroup

	for _, inputPath := range args {
		absInputPath, err := filepath.Abs(inputPath)
		if err != nil {
			log.Printf("Error resolving path %s: %v", inputPath, err)
			continue // Log and continue with the next argument
		}

		info, err := os.Stat(absInputPath)
		if err != nil {
			log.Printf("Error accessing path %s: %v", absInputPath, err)
			continue // Log and continue
		}

		if info.IsDir() {
			// Pass the WaitGroup to processDirectory
			// Note: processDirectory itself doesn't return an error to run
			processDirectory(extractor, absInputPath, recursive, &wg)
		} else {
			// Increment counter and launch goroutine for files directly specified
			wg.Add(1)
			go processFile(extractor, absInputPath, &wg)
		}
	}

	wg.Wait() // Wait for all goroutines to finish
	fmt.Println("--- All processing finished ---")
	return nil // Indicate success
}

// processDirectory now accepts a WaitGroup pointer
func processDirectory(extractor *word_extractor.WordExtractor, dirPath string, recursive bool, wg *sync.WaitGroup) {
	err := filepath.Walk(dirPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Printf("Error accessing path %q: %v\n", path, err)
			return err // Propagate the error to stop walking if needed, or return nil to continue
		}

		if info.IsDir() && !recursive && path != dirPath {
			return filepath.SkipDir
		}

		if !info.IsDir() && (strings.HasSuffix(strings.ToLower(info.Name()), ".docx") || strings.HasSuffix(strings.ToLower(info.Name()), ".doc")) {
			// Increment counter and launch goroutine for files found during walk
			wg.Add(1)
			go processFile(extractor, path, wg) // Pass the absolute path found by Walk
		}
		return nil
	})
	if err != nil {
		// Log the error from Walk itself, but don't necessarily exit
		log.Printf("Error walking directory %s: %v", dirPath, err)
	}
}

// processFile now accepts a WaitGroup pointer
func processFile(extractor *word_extractor.WordExtractor, filePath string, wg *sync.WaitGroup) {
	defer wg.Done() // Ensure Done is called when the function exits

	fmt.Printf("--- Processing file: %s ---\n", filePath)

	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		log.Printf("File does not exist at path: %s", filePath)
		return // Skip this file
	}

	doc, err := extractor.Extract(filePath)
	if err != nil {
		log.Printf("Error extracting content from %s: %v", filePath, err)
		return // Skip this file
	}

	// Print the extracted content
	fmt.Println("Extracted Headers:")
	headers := doc.GetHeaders(nil)
	fmt.Println(headers)

	fmt.Println("\nExtracted Body:")
	body := doc.GetBody(nil)
	fmt.Println(body)

	fmt.Printf("--- Finished processing file: %s ---\n\n", filePath)
}
