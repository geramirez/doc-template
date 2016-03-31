package main

import (
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

type docxTest struct {
	fixture string
	content string
	err     error
}

var readDocTests = []docxTest{
	//  Check that reading a document works
	{fixture: "fixtures/test.docx", content: "This is a test document", err: nil},
}

func TestReadDocxFile(t *testing.T) {
	for _, example := range readDocTests {
		actualData, actualErr := ReadDocxFile(example.fixture)
		assert.Equal(t, example.err, actualErr)
		if actualErr == nil {
			assert.Contains(t, *(actualData.content), example.content)
		}
	}
}

var writeDocTests = []docxTest{
	//  Check that writing a document works
	{fixture: "fixtures/test.docx", content: "This is an addition", err: nil},
}

func TestWriteToFile(t *testing.T) {
	for _, example := range writeDocTests {
		exportTempDir, _ := ioutil.TempDir("", "exports")
		// Overwrite content
		data, _ := ReadDocxFile(example.fixture)
		newstring := strings.Replace(*(data.content), "This is a test document", example.content, -1)
		data.content = &newstring
		newFilePath := filepath.Join(exportTempDir, "test.docx")
		data.WriteToFile(newFilePath)
		// Check content
		actualData, _ := ReadDocxFile(newFilePath)
		assert.Contains(t, *(actualData.content), example.content)
		os.RemoveAll(exportTempDir)
	}

}
