package docTemp

import (
	"bytes"
	"errors"
	"path/filepath"
	"strings"
	"text/template"

	"github.com/geramirez/doc-template/docx"
)

// Document interface is a combintation of methods use for generic data files
type Document interface {
	ReadFile(string) error
	UpdateConent(string)
	GetContent() string
	WriteToFile(string, string) error
	Close() error
}

// DocTemplate struct combines data and methods from both the Document interface
// and golang's templating library
type DocTemplate struct {
	Template *template.Template
	Document Document
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
	return &DocTemplate{Document: document, Template: template.New("docTemp")}, nil
}

// Execute func runs the template and sends the output to the export path
func (docTemplate *DocTemplate) Execute(exportPath string, data interface{}) error {
	buf := new(bytes.Buffer)
	err := docTemplate.Template.Execute(buf, data)
	if err != nil {
		return err
	}
	err = docTemplate.Document.WriteToFile(exportPath, buf.String())
	return err
}

// AddFunctions adds functions to the template
func (docTemplate *DocTemplate) AddFunctions(funcMap template.FuncMap) {
	docTemplate.Template = docTemplate.Template.Funcs(funcMap)
}

// Parse parses the template
func (docTemplate *DocTemplate) Parse() {
	docTemplate.Template.Parse(docTemplate.Document.GetContent())
}

func main() {
	funcMap := template.FuncMap{"title": strings.Title}
	docTemp, _ := GetTemplate("docx/fixtures/test.docx")
	docTemp.AddFunctions(funcMap)
	docTemp.Parse()
	docTemp.Execute("test.docx", nil)
}
