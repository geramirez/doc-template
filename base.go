package main

import (
	"errors"
	"html/template"
	"path/filepath"

	"github.com/geramirez/doc-template/docx"
)

// Document interface is a combintation of methods use for generic data files
type Document interface {
	ReadFile(string) error
	UpdateConent(string)
	GetContent() string
	WriteToFile(string) error
	Close() error
}

// DocTemplate struct combines data and methods from both the Document interface
// and golang's templating library
type DocTemplate struct {
	Document Document
	Template *template.Template
}

// GetTemplate uses the file extension to determin the correct document struct to use
func GetTemplate(filePath string) (*DocTemplate, error) {
	var document Document
	switch filepath.Ext(filePath) {
	case ".docx":
		document = new(docx.Docx)
	default:
		return nil, errors.New("Unsupported Document Type")
	}
	document.ReadFile(filePath)
	return &DocTemplate{Document: document, Template: template.New("newTemplate")}, nil
}
