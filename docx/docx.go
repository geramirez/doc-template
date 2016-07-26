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
			//WORKS on image placeholder text writer.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" mc:Ignorable="w14"><w:body><w:p><w:pPr><w:pStyle w:val="Body"/><w:bidi w:val="0"/></w:pPr><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:drawing><wp:anchor distT="152400" distB="152400" distL="152400" distR="152400" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>1720850</wp:posOffset></wp:positionH><wp:positionV relativeFrom="line"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="4318000" cy="787400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapThrough wrapText="bothSides" distL="152400" distR="152400"><wp:wrapPolygon edited="1"></wp:wrapPolygon></wp:wrapThrough><wp:docPr id="1073741825" name="officeArt object"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1073741825" name="pasted-image.png"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId4"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4318000" cy="787400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln><a:effectLst/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:footerReference w:type="default" r:id="rId6"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="864"/><w:bidi w:val="0"/></w:sectPr></w:body></w:document>`))
			/*writer.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14"><w:body><w:p><w:pPr><w:pStyle w:val="BodyA"/><w:shd w:fill="FFFFFF" w:val="clear"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:t>Image:</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="BodyA"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr></w:r></w:p><w:p><w:pPr><w:pStyle w:val="BodyA"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr><w:drawing><wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="2"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV><wp:extent cx="2086610" cy="380365"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="largest"/><wp:docPr id="1" name="Image1" descr=""></wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Image1" descr=""></pic:cNvPr><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="2086610" cy="380365"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:header="720" w:top="1440" w:footer="864" w:bottom="1440" w:gutter="0"/><w:pgNumType w:fmt="decimal"/><w:formProt w:val="false"/><w:textDirection w:val="lrTb"/><w:bidi/><w:docGrid w:type="default" w:linePitch="240" w:charSpace="4294961151"/></w:sectPr></w:body></w:document>
`))*/
			writer.Write([]byte(data))
			//writer.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" mc:Ignorable="w14"><w:body><w:p><w:pPr><w:pStyle w:val="Body"/><w:bidi w:val="0"/></w:pPr><w:r><w:drawing><wp:anchor distT="152400" distB="152400" distL="152400" distR="152400" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>806450</wp:posOffset></wp:positionH><wp:positionV relativeFrom="line"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="4318000" cy="787400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapThrough wrapText="bothSides" distL="152400" distR="152400"><wp:wrapPolygon edited="1"><wp:start x="445" y="0"/><wp:lineTo x="889" y="697"/><wp:lineTo x="1080" y="11148"/><wp:lineTo x="1080" y="16374"/><wp:lineTo x="762" y="20555"/><wp:lineTo x="254" y="19858"/><wp:lineTo x="127" y="16026"/><wp:lineTo x="127" y="10800"/><wp:lineTo x="254" y="5574"/><wp:lineTo x="953" y="12194"/><wp:lineTo x="254" y="11845"/><wp:lineTo x="318" y="15677"/><wp:lineTo x="953" y="15329"/><wp:lineTo x="953" y="12194"/><wp:lineTo x="254" y="5574"/><wp:lineTo x="381" y="348"/><wp:lineTo x="445" y="0"/><wp:lineTo x="1715" y="0"/><wp:lineTo x="2224" y="1045"/><wp:lineTo x="2351" y="3135"/><wp:lineTo x="2351" y="8710"/><wp:lineTo x="2096" y="20555"/><wp:lineTo x="1588" y="20206"/><wp:lineTo x="1398" y="8361"/><wp:lineTo x="1398" y="3135"/><wp:lineTo x="1525" y="1879"/><wp:lineTo x="2160" y="4181"/><wp:lineTo x="1525" y="4181"/><wp:lineTo x="1588" y="7665"/><wp:lineTo x="2224" y="7665"/><wp:lineTo x="2160" y="4181"/><wp:lineTo x="1525" y="1879"/><wp:lineTo x="1715" y="0"/><wp:lineTo x="3049" y="0"/><wp:lineTo x="3494" y="697"/><wp:lineTo x="3685" y="8710"/><wp:lineTo x="3685" y="13935"/><wp:lineTo x="3431" y="20206"/><wp:lineTo x="2922" y="20206"/><wp:lineTo x="2732" y="13587"/><wp:lineTo x="2732" y="8013"/><wp:lineTo x="2859" y="2889"/><wp:lineTo x="3494" y="9406"/><wp:lineTo x="2859" y="9058"/><wp:lineTo x="2859" y="12890"/><wp:lineTo x="3558" y="12542"/><wp:lineTo x="3494" y="9406"/><wp:lineTo x="2859" y="2889"/><wp:lineTo x="2922" y="348"/><wp:lineTo x="3049" y="0"/><wp:lineTo x="4892" y="0"/><wp:lineTo x="4892" y="2439"/><wp:lineTo x="5464" y="2920"/><wp:lineTo x="5464" y="4181"/><wp:lineTo x="4765" y="4529"/><wp:lineTo x="4828" y="15677"/><wp:lineTo x="5591" y="15677"/><wp:lineTo x="5591" y="4529"/><wp:lineTo x="5464" y="4181"/><wp:lineTo x="5464" y="2920"/><wp:lineTo x="5718" y="3135"/><wp:lineTo x="5972" y="5226"/><wp:lineTo x="5845" y="16374"/><wp:lineTo x="5591" y="17419"/><wp:lineTo x="4638" y="16723"/><wp:lineTo x="4447" y="15329"/><wp:lineTo x="4574" y="3484"/><wp:lineTo x="4892" y="2439"/><wp:lineTo x="4892" y="0"/><wp:lineTo x="6861" y="0"/><wp:lineTo x="6861" y="6619"/><wp:lineTo x="7560" y="7316"/><wp:lineTo x="7560" y="16723"/><wp:lineTo x="6671" y="16723"/><wp:lineTo x="6671" y="21600"/><wp:lineTo x="6416" y="21600"/><wp:lineTo x="6544" y="6968"/><wp:lineTo x="6734" y="6968"/><wp:lineTo x="7306" y="8361"/><wp:lineTo x="6734" y="8361"/><wp:lineTo x="6798" y="16026"/><wp:lineTo x="7433" y="15677"/><wp:lineTo x="7369" y="8361"/><wp:lineTo x="7306" y="8361"/><wp:lineTo x="6734" y="6968"/><wp:lineTo x="6861" y="6968"/><wp:lineTo x="6861" y="6619"/><wp:lineTo x="6861" y="0"/><wp:lineTo x="8513" y="0"/><wp:lineTo x="8513" y="6619"/><wp:lineTo x="9021" y="7083"/><wp:lineTo x="9021" y="8361"/><wp:lineTo x="8449" y="8710"/><wp:lineTo x="8449" y="11148"/><wp:lineTo x="9212" y="11148"/><wp:lineTo x="9148" y="8361"/><wp:lineTo x="9021" y="8361"/><wp:lineTo x="9021" y="7083"/><wp:lineTo x="9275" y="7316"/><wp:lineTo x="9466" y="8710"/><wp:lineTo x="9466" y="12194"/><wp:lineTo x="8386" y="12542"/><wp:lineTo x="8513" y="16026"/><wp:lineTo x="9148" y="15677"/><wp:lineTo x="9212" y="13935"/><wp:lineTo x="9466" y="13935"/><wp:lineTo x="9275" y="17071"/><wp:lineTo x="8322" y="16723"/><wp:lineTo x="8322" y="7316"/><wp:lineTo x="8513" y="6619"/><wp:lineTo x="8513" y="0"/><wp:lineTo x="10355" y="0"/><wp:lineTo x="10355" y="6619"/><wp:lineTo x="11054" y="7316"/><wp:lineTo x="11181" y="17419"/><wp:lineTo x="10927" y="17419"/><wp:lineTo x="10800" y="8013"/><wp:lineTo x="10165" y="8710"/><wp:lineTo x="10165" y="17419"/><wp:lineTo x="9911" y="17419"/><wp:lineTo x="10038" y="6968"/><wp:lineTo x="10355" y="6968"/><wp:lineTo x="10355" y="6619"/><wp:lineTo x="10355" y="0"/><wp:lineTo x="12134" y="0"/><wp:lineTo x="12134" y="2439"/><wp:lineTo x="12960" y="3135"/><wp:lineTo x="13214" y="6619"/><wp:lineTo x="12896" y="6271"/><wp:lineTo x="12769" y="4181"/><wp:lineTo x="12007" y="4529"/><wp:lineTo x="12071" y="15677"/><wp:lineTo x="12769" y="15677"/><wp:lineTo x="12960" y="13239"/><wp:lineTo x="13151" y="13239"/><wp:lineTo x="13024" y="16723"/><wp:lineTo x="11880" y="16723"/><wp:lineTo x="11689" y="15329"/><wp:lineTo x="11816" y="3484"/><wp:lineTo x="12134" y="2439"/><wp:lineTo x="12134" y="0"/><wp:lineTo x="13976" y="0"/><wp:lineTo x="13976" y="6619"/><wp:lineTo x="14485" y="7083"/><wp:lineTo x="14485" y="8361"/><wp:lineTo x="13913" y="8710"/><wp:lineTo x="13976" y="16026"/><wp:lineTo x="14612" y="15677"/><wp:lineTo x="14612" y="8361"/><wp:lineTo x="14485" y="8361"/><wp:lineTo x="14485" y="7083"/><wp:lineTo x="14739" y="7316"/><wp:lineTo x="14929" y="8710"/><wp:lineTo x="14802" y="16723"/><wp:lineTo x="13786" y="16723"/><wp:lineTo x="13786" y="7316"/><wp:lineTo x="13976" y="6619"/><wp:lineTo x="13976" y="0"/><wp:lineTo x="15819" y="0"/><wp:lineTo x="15819" y="6619"/><wp:lineTo x="16518" y="7316"/><wp:lineTo x="16645" y="17419"/><wp:lineTo x="16391" y="17419"/><wp:lineTo x="16264" y="8013"/><wp:lineTo x="15628" y="8710"/><wp:lineTo x="15628" y="17419"/><wp:lineTo x="15374" y="17419"/><wp:lineTo x="15501" y="6968"/><wp:lineTo x="15819" y="6968"/><wp:lineTo x="15819" y="6619"/><wp:lineTo x="15819" y="0"/><wp:lineTo x="17280" y="0"/><wp:lineTo x="17280" y="4181"/><wp:lineTo x="17534" y="4181"/><wp:lineTo x="17534" y="6968"/><wp:lineTo x="17979" y="7316"/><wp:lineTo x="18042" y="8013"/><wp:lineTo x="17534" y="8013"/><wp:lineTo x="17661" y="16026"/><wp:lineTo x="17979" y="16026"/><wp:lineTo x="17979" y="17419"/><wp:lineTo x="17344" y="16374"/><wp:lineTo x="17280" y="8013"/><wp:lineTo x="16962" y="8013"/><wp:lineTo x="16962" y="6968"/><wp:lineTo x="17280" y="6968"/><wp:lineTo x="17280" y="4181"/><wp:lineTo x="17280" y="0"/><wp:lineTo x="18868" y="0"/><wp:lineTo x="18868" y="6619"/><wp:lineTo x="19186" y="6619"/><wp:lineTo x="19249" y="8013"/><wp:lineTo x="18741" y="8710"/><wp:lineTo x="18678" y="17419"/><wp:lineTo x="18424" y="17419"/><wp:lineTo x="18551" y="6968"/><wp:lineTo x="18868" y="6968"/><wp:lineTo x="18868" y="6619"/><wp:lineTo x="18868" y="0"/><wp:lineTo x="19885" y="0"/><wp:lineTo x="19885" y="6619"/><wp:lineTo x="20647" y="7316"/><wp:lineTo x="20647" y="16723"/><wp:lineTo x="19631" y="16723"/><wp:lineTo x="19631" y="7665"/><wp:lineTo x="19821" y="6882"/><wp:lineTo x="20393" y="8361"/><wp:lineTo x="19821" y="8361"/><wp:lineTo x="19885" y="16026"/><wp:lineTo x="20456" y="16026"/><wp:lineTo x="20393" y="8361"/><wp:lineTo x="19821" y="6882"/><wp:lineTo x="19885" y="6619"/><wp:lineTo x="19885" y="0"/><wp:lineTo x="445" y="0"/></wp:wrapPolygon></wp:wrapThrough><wp:docPr id="1073741825" name="officeArt object"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1073741825" name="pasted-image.png"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId4"><a:extLst/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4318000" cy="787400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln><a:effectLst/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:footerReference w:type="default" r:id="rId6"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="864"/><w:bidi w:val="0"/></w:sectPr></w:body></w:document>`))
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
