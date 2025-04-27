package main

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// Helper function to create a temporary directory with a dummy file for testing
func createTestDir(t *testing.T) (string, func()) {
	t.Helper()
	tmpDir, err := os.MkdirTemp("", "testrundir")
	if err != nil {
		t.Fatalf("Failed to create temp dir: %v", err)
	}

	// Create a dummy file inside
	dummyFilePath := filepath.Join(tmpDir, "dummy.doc")
	err = os.WriteFile(dummyFilePath, []byte("dummy content"), 0644)
	if err != nil {
		os.RemoveAll(tmpDir) // Clean up if file creation fails
		t.Fatalf("Failed to create dummy file: %v", err)
	}

	// Create a subdirectory with another dummy file for recursive test
	subDir := filepath.Join(tmpDir, "subdir")
	err = os.Mkdir(subDir, 0755)
	if err != nil {
		os.RemoveAll(tmpDir)
		t.Fatalf("Failed to create sub dir: %v", err)
	}
	dummySubFilePath := filepath.Join(subDir, "subdummy.docx")
	err = os.WriteFile(dummySubFilePath, []byte("sub dummy content"), 0644)
	if err != nil {
		os.RemoveAll(tmpDir)
		t.Fatalf("Failed to create sub dummy file: %v", err)
	}

	cleanup := func() {
		os.RemoveAll(tmpDir)
	}
	return tmpDir, cleanup
}

func TestRun_NoArgs(t *testing.T) {
	err := run([]string{}, false)
	if err == nil {
		t.Fatal("Expected an error when no arguments are provided, but got nil")
	}
	expectedErrMsg := "usage: go-word-extractor [-r] <file_or_dir1> [file_or_dir2]"
	if !strings.Contains(err.Error(), expectedErrMsg) {
		t.Errorf("Expected error message to contain '%s', but got '%s'", expectedErrMsg, err.Error())
	}
}

func TestRun_ValidFile(t *testing.T) {
	// Assuming tests/data/test01.docx exists and is readable
	testFilePath := filepath.Join("tests", "data", "test01.docx")
	if _, err := os.Stat(testFilePath); os.IsNotExist(err) {
		t.Skipf("Skipping test: Test file %s does not exist", testFilePath)
	}

	err := run([]string{testFilePath}, false)
	if err != nil {
		t.Errorf("Expected no error when processing a valid file, but got: %v", err)
	}
	// Note: This test doesn't verify the *output*, only that the run completes without returning an error.
}

func TestRun_ValidDirNonRecursive(t *testing.T) {
	tmpDir, cleanup := createTestDir(t)
	defer cleanup()

	err := run([]string{tmpDir}, false) // Should process dummy.doc, skip subdir/subdummy.docx
	if err != nil {
		t.Errorf("Expected no error when processing a valid directory non-recursively, but got: %v", err)
	}
	// Further checks could involve capturing/mocking log output if needed
}

func TestRun_ValidDirRecursive(t *testing.T) {
	tmpDir, cleanup := createTestDir(t)
	defer cleanup()

	err := run([]string{tmpDir}, true) // Should process dummy.doc AND subdir/subdummy.docx
	if err != nil {
		t.Errorf("Expected no error when processing a valid directory recursively, but got: %v", err)
	}
	// Further checks could involve capturing/mocking log output if needed
}

func TestRun_InvalidPath(t *testing.T) {
	invalidPath := filepath.Join("non", "existent", "path", "file.doc")
	err := run([]string{invalidPath}, false)
	if err != nil {
		t.Errorf("Expected no error returned from run() for an invalid path (error should be logged), but got: %v", err)
	}
	// Note: This assumes the desired behavior is to log the error and continue,
	// so run() itself returns nil unless *all* inputs fail in a way that stops processing early.
}

func TestRun_MixedPaths(t *testing.T) {
	testFilePath := filepath.Join("tests", "data", "test01.docx")
	if _, err := os.Stat(testFilePath); os.IsNotExist(err) {
		t.Skipf("Skipping test: Test file %s does not exist", testFilePath)
	}
	invalidPath := filepath.Join("non", "existent", "path", "file.doc")

	err := run([]string{testFilePath, invalidPath}, false)
	if err != nil {
		t.Errorf("Expected no error returned from run() for mixed valid/invalid paths, but got: %v", err)
	}
}
