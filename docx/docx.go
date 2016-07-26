package docx

import (
	"archive/zip"
	"errors"
	"io"
	"io/ioutil"
	"log"
	"os"
	"net/http"
	"fmt"
	"path/filepath"
)

// Docx struct that contains data from a docx
type Docx struct {
	zipReader *zip.ReadCloser
	content   string
	pictures  []*Picture
}

// ReadFile func reads a docx file
func (d *Docx) ReadFile(path string) error {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return errors.New("Cannot Open File")
	}
	content, err := readText(reader.File)
	if err != nil {
		return errors.New("Cannot Read File")
	}
	d.zipReader = reader
	if content == "" {
		return errors.New("File has no content")
	}
	d.content = cleanText(content)
	log.Printf("Read File `%s`", path)
	return nil
}

// UpdateContent updates the content string
func (d *Docx) UpdateContent(newContent string) {
	d.content = newContent
}

// GetContent returns the string content
func (d *Docx) GetContent() string {
	return d.content
}

// WriteToFile writes the changes to a new file
func (d *Docx) WriteToFile(path string, data string) error {
	var target *os.File
	target, err := os.Create(path)
	if err != nil {
		return err
	}
	defer target.Close()
	err = d.write(target, data)
	if err != nil {
		return err
	}
	log.Printf("Exporting data to %s", path)
	return nil
}

// Close the document
func (d *Docx) Close() error {
	return d.zipReader.Close()
}

func (d *Docx) write(ioWriter io.Writer, data string) error {
	var err error
	// Reformat string, for some reason the first char is converted to &lt;
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
			writer.Write([]byte(data))
		} else if file.Name == "word/_rels/document.xml.rels"{
			writer.Write([]byte(testRelsContent))
		} else {
			writer.Write(streamToByte(readCloser))
		}
		readCloser.Close()
	}
	for _, picture := range(d.pictures) {
		var writer io.Writer
		writer, err := w.Create(filepath.Join("word", "media", picture.name))
		if err != nil {
			return err
		}
		writer.Write(picture.data)
	}
	w.Close()
	return err
}

func (d *Docx) UploadImage(path string) (*Picture, error){
	// Try to read the image.
	imageFile, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	// Verify the file is actually an image.
	fileType := http.DetectContentType(imageFile)
	switch fileType {
	case "image/gif", "image/png", "image/jpeg":
		// Good. Will not fall through.
	default:
		return nil, fmt.Errorf("Unsupported Image type. Type: %s", fileType)
	}

	// Get the checksum of the new image.
	newImageCheckSum := calculateChecksum(imageFile)

	// Check to see if the file already exists.
	// If it does, just return the existing Picture handle.
	for _, file := range d.zipReader.File {
		readCloser ,err := file.Open()
		if err != nil {
			return nil, err
		}
		fileBytes := streamToByte(readCloser)
		fileCheckSum := calculateChecksum(fileBytes)
		if newImageCheckSum == fileCheckSum {
			// TODO: get the picture struct and return
		}
	}

	// TODO go through d.images and check the checksum

	pic := NewPicture(imageFile, "image1.png", fileType)
	d.pictures = append(d.pictures, pic)

	return pic, nil
}
