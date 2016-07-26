package docx

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

func TestReadFile(t *testing.T) {
	for _, example := range readDocTests {
		actualDoc := new(Docx)
		actualErr := actualDoc.ReadFile(example.fixture)
		assert.Equal(t, example.err, actualErr)
		if actualErr == nil {
			assert.Contains(t, actualDoc.content, example.content)
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
		actualDoc := new(Docx)
		actualDoc.ReadFile(example.fixture)
		currentContent := actualDoc.GetContent()
		actualDoc.UpdateContent(strings.Replace(currentContent, "This is a test document", example.content, -1))
		newFilePath := filepath.Join(exportTempDir, "test.docx")
		actualDoc.WriteToFile(newFilePath, actualDoc.GetContent())
		// Check content
		newActualDoc := new(Docx)
		newActualDoc.ReadFile(newFilePath)
		assert.Contains(t, newActualDoc.GetContent(), example.content)
		os.RemoveAll(exportTempDir)
	}

}


type mediaDocxTest struct {
	fixture string
	image string
	expected string
}

var mediaDocxTests = []mediaDocxTest {
	{
		fixture: filepath.Join("fixtures", "picture_fixtures", "empty.docx"),
		image: filepath.Join("fixtures", "picture_fixtures", "oclogo.png"),
	},
}

func TestUploadImage(t *testing.T){
	for _, test := range mediaDocxTests {
		exportTempDir, _ := ioutil.TempDir("", "exports")
		actualDoc := new(Docx)
		assert.Nil(t, actualDoc.ReadFile(test.fixture))
		_, err := actualDoc.UploadImage(test.image)
		assert.Nil(t, err)
		//t.Log(actualDoc.GetContent())
		actualDoc.UpdateContent(testDocumentContent)
		//t.Log(actualDoc.GetContent())
		exportTempDir = ""
		newFilePath := filepath.Join(exportTempDir, "test.docx")
		actualDoc.WriteToFile(newFilePath, actualDoc.GetContent())
		actualDoc.Close()
		// Check content
		newActualDoc := new(Docx)
		newActualDoc.ReadFile(newFilePath)
		for _, file := range(newActualDoc.zipReader.File) {
			// t.Log(file.Name)
			if file.Name == "word/media/image1.png" {
				reader, _ := file.Open()
				ioutil.WriteFile("image1.png", streamToByte(reader), 0644)
			}
		}
		//t.Log("Finished")
		assert.Equal(t, newActualDoc.GetContent(), actualDoc.GetContent())
		os.RemoveAll(exportTempDir)
		newActualDoc.Close()


	}
}
