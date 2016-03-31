package main

import "html/template"

// Document interface is a combintation of methods use for generic data files
type Document interface {
	ReadFile(string)
	UpdateConent(string)
	GetContent(string)
	WriteToFile(string) error
	Close() error
}

// DocTemplate struct combines data and methods from both the Document interface
// and golang's templating library
type DocTemplate struct {
	Document *Document
	Template template.Template
}
