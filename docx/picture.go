package docx

import "fmt"

func NewPicture(data []byte, name string, mime string) *Picture {
	// Create a copy of the data
	tmp := make([]byte, len(data))
	copy(tmp, data)
	return &Picture{checksum: 0, name: name, data:tmp, mime:mime}
}

type Picture struct {
	checksum uint32
	name string
	data []byte
	mime string
}

func (p Picture) GetRepresentation() string {
	rep := fmt.Sprintf(`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <pic:nvPicPr>
    <pic:cNvPr id="1" name="%s"/>
    <pic:cNvPicPr/>
  </pic:nvPicPr>
  <pic:blipFill>
    <a:blip r:embed="rId7"/>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </pic:blipFill>
  <pic:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="859536" cy="343814"/>
    </a:xfrm>
    <a:prstGeom prst="rect"/>
  </pic:spPr>
</pic:pic>`, p.name)
	return rep
}