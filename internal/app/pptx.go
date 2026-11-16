package app

import (
	"archive/zip"
	"bytes"
	"fmt"
	"net/http"
	"net/url"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// PowerPoint export. The deck is built from the text form of the entry
// file (text.go): one slide per .slide of a slide deck, or per h1/h2 of any
// other page, with a title and editable text. The page's own design is not
// carried over. The same code is in cloudflare/src/pptx.ts.

type pptSlide struct {
	title string
	paras []pptPara
}

type pptPara struct {
	text         string
	bullet, bold bool
}

var (
	rePptSlide   = regexp.MustCompile(`(?m)^\[Slide \d+\]$`)
	rePptSection = regexp.MustCompile(`(?m)^#{1,2} `)
	rePptHeading = regexp.MustCompile(`^#{1,6} +`)
	rePptCtl     = regexp.MustCompile(`[\x00-\x08\x0b\x0c\x0e-\x1f]`)
)

// pptSlides cuts the text form of a page into slides.
func pptSlides(text string) []pptSlide {
	var chunks []string
	if rePptSlide.MatchString(text) {
		chunks = rePptSlide.Split(text, -1)
	} else {
		at := 0
		for _, m := range rePptSection.FindAllStringIndex(text, -1) {
			chunks = append(chunks, text[at:m[0]])
			at = m[0]
		}
		chunks = append(chunks, text[at:])
	}
	var out []pptSlide
	for _, c := range chunks {
		var s pptSlide
		for _, line := range strings.Split(c, "\n") {
			line = strings.TrimSpace(line)
			h := rePptHeading.FindString(line)
			switch {
			case line == "" || line == "---" || strings.HasPrefix(line, "```"):
			case h != "" && s.title == "":
				s.title = line[len(h):]
			case h != "":
				s.paras = append(s.paras, pptPara{text: line[len(h):], bold: true})
			case strings.HasPrefix(line, "- "):
				s.paras = append(s.paras, pptPara{text: line[2:], bullet: true})
			default:
				s.paras = append(s.paras, pptPara{text: line})
			}
		}
		if s.title == "" && len(s.paras) > 0 {
			s.title, s.paras = s.paras[0].text, s.paras[1:]
		}
		if s.title != "" {
			out = append(out, s)
		}
	}
	return out
}

func pptEsc(s string) string {
	s = rePptCtl.ReplaceAllString(s, "")
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;").Replace(s)
}

const (
	pptHead = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	pptNS   = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`
	pptRel  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
	pptType = "application/vnd.openxmlformats-officedocument."
	pptTree = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`

	pptTheme = pptHead + `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="9artifacts"><a:themeElements>` +
		`<a:clrScheme name="9artifacts"><a:dk1><a:srgbClr val="1A1A1A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44403C"/></a:dk2><a:lt2><a:srgbClr val="F5F5F4"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="C2410C"/></a:accent1><a:accent2><a:srgbClr val="0F766E"/></a:accent2><a:accent3><a:srgbClr val="1D4ED8"/></a:accent3><a:accent4><a:srgbClr val="A16207"/></a:accent4><a:accent5><a:srgbClr val="7E22CE"/></a:accent5><a:accent6><a:srgbClr val="BE123C"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="1D4ED8"/></a:hlink><a:folHlink><a:srgbClr val="7E22CE"/></a:folHlink></a:clrScheme>` +
		`<a:fontScheme name="9artifacts"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
		`<a:fmtScheme name="9artifacts"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
		`<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
		`<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
		`<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
		`</a:themeElements></a:theme>`

	// Title and body boxes on a 16:9 slide (12192000 × 6858000 EMU).
	pptTitleBox = `<a:xfrm><a:off x="685800" y="457200"/><a:ext cx="10820400" cy="1143000"/></a:xfrm>`
	pptBodyBox  = `<a:xfrm><a:off x="685800" y="1752600"/><a:ext cx="10820400" cy="4572000"/></a:xfrm>`

	pptMaster = pptHead + `<p:sldMaster ` + pptNS + `><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>` + pptTree +
		`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>` + pptTitleBox + `</p:spPr><p:txBody><a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>` +
		`<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>` + pptBodyBox + `</p:spPr><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>` +
		`</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
		`<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles>` +
		`<p:titleStyle><a:lvl1pPr algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/><a:defRPr sz="3600" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
		`<p:bodyStyle><a:lvl1pPr marL="0" indent="0" algn="l"><a:spcBef><a:spcPts val="800"/></a:spcBef><a:buNone/><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle>` +
		`<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`

	pptLayout = pptHead + `<p:sldLayout ` + pptNS + ` type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree>` + pptTree +
		`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>` +
		`<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>` +
		`</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
)

func pptRels(targets ...string) string {
	b := pptHead + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
	for i := 0; i+1 < len(targets); i += 2 {
		b += fmt.Sprintf(`<Relationship Id="rId%d" Type="%s%s" Target="%s"/>`, i/2+1, pptRel, targets[i], targets[i+1])
	}
	return b + `</Relationships>`
}

func pptSlideXML(s pptSlide) string {
	// Long slides get smaller text so they still fit the body box.
	size := 2000
	switch n := len(s.paras); {
	case n > 16:
		size = 1100
	case n > 9:
		size = 1400
	}
	var body strings.Builder
	for _, p := range s.paras {
		body.WriteString(`<a:p>`)
		if p.bullet {
			body.WriteString(`<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>`)
		}
		bold := ""
		if p.bold {
			bold = ` b="1"`
		}
		fmt.Fprintf(&body, `<a:r><a:rPr lang="en-US" sz="%d"%s/><a:t>%s</a:t></a:r></a:p>`, size, bold, pptEsc(p.text))
	}
	if len(s.paras) == 0 {
		body.WriteString(`<a:p><a:endParaRPr lang="en-US"/></a:p>`)
	}
	return pptHead + `<p:sld ` + pptNS + `><p:cSld><p:spTree>` + pptTree +
		`<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>` + pptEsc(s.title) + `</a:t></a:r></a:p></p:txBody></p:sp>` +
		`<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>` + body.String() + `</p:txBody></p:sp>` +
		`</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
}

// pptParts returns the files of a .pptx package in the order they are written.
func pptParts(slides []pptSlide) [][2]string {
	types := pptHead + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/ppt/presentation.xml" ContentType="` + pptType + `presentationml.presentation.main+xml"/>` +
		`<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="` + pptType + `presentationml.slideMaster+xml"/>` +
		`<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="` + pptType + `presentationml.slideLayout+xml"/>` +
		`<Override PartName="/ppt/theme/theme1.xml" ContentType="` + pptType + `theme+xml"/>`
	ids := ""
	rels := []string{"slideMaster", "slideMasters/slideMaster1.xml", "theme", "theme/theme1.xml"}
	var parts [][2]string
	for i, s := range slides {
		name := fmt.Sprintf("slides/slide%d.xml", i+1)
		types += `<Override PartName="/ppt/` + name + `" ContentType="` + pptType + `presentationml.slide+xml"/>`
		ids += fmt.Sprintf(`<p:sldId id="%d" r:id="rId%d"/>`, 256+i, i+3)
		rels = append(rels, "slide", name)
		parts = append(parts,
			[2]string{"ppt/" + name, pptSlideXML(s)},
			[2]string{fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i+1), pptRels("slideLayout", "../slideLayouts/slideLayout1.xml")})
	}
	pres := pptHead + `<p:presentation ` + pptNS + `><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>` + ids +
		`</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`
	return append([][2]string{
		{"[Content_Types].xml", types + `</Types>`},
		{"_rels/.rels", pptRels("officeDocument", "ppt/presentation.xml")},
		{"ppt/presentation.xml", pres},
		{"ppt/_rels/presentation.xml.rels", pptRels(rels...)},
		{"ppt/slideMasters/slideMaster1.xml", pptMaster},
		{"ppt/slideMasters/_rels/slideMaster1.xml.rels", pptRels("slideLayout", "../slideLayouts/slideLayout1.xml", "theme", "../theme/theme1.xml")},
		{"ppt/slideLayouts/slideLayout1.xml", pptLayout},
		{"ppt/slideLayouts/_rels/slideLayout1.xml.rels", pptRels("slideMaster", "../slideMasters/slideMaster1.xml")},
		{"ppt/theme/theme1.xml", pptTheme},
	}, parts...)
}

// handlePPTX serves GET /api/artifacts/{id}/pptx: the entry file as an
// editable PowerPoint deck.
func (s *Server) handlePPTX(w http.ResponseWriter, r *http.Request) {
	id := r.PathValue("id")
	n, _ := strconv.Atoi(r.URL.Query().Get("v"))
	var v *Version
	var title string
	s.store.Read(func(st *State) {
		if a := st.Artifacts[id]; a != nil {
			if x := a.Version(n); x != nil {
				c := *x
				v, title = &c, a.Title
			}
		}
	})
	if v == nil {
		writeJSON(w, 404, map[string]string{"error": "artifact or version not found"})
		return
	}
	if !searchable[strings.ToLower(path.Ext(v.Entry))] {
		writeJSON(w, 415, map[string]string{"error": v.Entry + " is not text, so it has no slides to export"})
		return
	}
	_, body, _ := strings.Cut(s.artifactText(id, title, v, nil), "\n---\n\n")
	slides := pptSlides(body)
	if len(slides) == 0 {
		slides = []pptSlide{{title: title}}
	}
	var buf bytes.Buffer
	z := zip.NewWriter(&buf)
	for _, p := range pptParts(slides) {
		f, _ := z.Create(p[0])
		_, _ = f.Write([]byte(p[1]))
	}
	_ = z.Close()
	name := strings.Trim(regexp.MustCompile(`[^A-Za-z0-9._-]+`).ReplaceAllString(title, "-"), "-.")
	if name == "" {
		name = id
	}
	w.Header().Set("Content-Type", pptType+"presentationml.presentation")
	w.Header().Set("Content-Disposition", `attachment; filename="`+name+`.pptx"; filename*=UTF-8''`+url.PathEscape(title+".pptx"))
	w.Header().Set("Cache-Control", "no-store")
	_, _ = w.Write(buf.Bytes())
}
