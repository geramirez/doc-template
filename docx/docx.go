package docx

import (
	"archive/zip"
	"io"
	"os"
)

// Docx struct that contains data from a docx
type Docx struct {
	zipReader *zip.ReadCloser
	content   string
}

// ReadFile func reads a docx file
func (d *Docx) ReadFile(path string) error {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return err
	}
	content, err := readText(reader.File)
	if err != nil {
		return err
	}
	d.zipReader = reader
	d.content = content
	return nil
}

// UpdateConent updates the content string
func (d *Docx) UpdateConent(newContent string) {
	d.content = newContent
}

// GetContent returns the string content
func (d *Docx) GetContent() string {
	return d.content
}

// WriteToFile writes the changes to a new file
func (d *Docx) WriteToFile(path string) error {
	var target *os.File
	target, err := os.Create(path)
	if err != nil {
		return err
	}
	defer target.Close()
	err = d.write(target)
	if err != nil {
		return err
	}
	return nil
}

// Closes the document
func (d *Docx) Close() error {
	return d.zipReader.Close()
}

func (d *Docx) write(ioWriter io.Writer) error {
	var err error
	w := zip.NewWriter(ioWriter)
	for _, file := range d.zipReader.File {
		var writer io.Writer
		var readCloser io.ReadCloser
		writer, err := w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		if file.Name == "word/document.xml" {
			writer.Write([]byte(d.content))
		} else {
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return err
}
