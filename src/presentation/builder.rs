use crate::parser::markdown::{MarkdownDocument, Slide, SlideElement};
use crate::presentation::templates::SlideTemplate;
use chrono::{DateTime, Utc};
use std::io::{Cursor, Write};
use uuid::Uuid;
use zip::ZipWriter;

use crate::Result;

pub struct PresentationBuilder {
    _template: SlideTemplate,
    slides: Vec<PptxSlide>,
    metadata: PresentationMetadata,
}

#[derive(Debug, Clone)]
struct PptxSlide {
    _id: String,
    title: Option<String>,
    content: Vec<PptxElement>,
}

#[derive(Debug, Clone)]
enum PptxElement {
    _Title(String),
    Text(String),
    BulletList(Vec<String>),
    NumberedList(Vec<String>),
    Code {
        _language: Option<String>,
        content: String,
    },
    Image {
        alt: String,
        _url: String,
    },
    Table {
        headers: Vec<String>,
        _rows: Vec<Vec<String>>,
    },
    Quote(String),
}

#[derive(Debug, Clone)]
struct PresentationMetadata {
    title: String,
    author: String,
    created: DateTime<Utc>,
    modified: DateTime<Utc>,
    slide_count: usize,
}

impl PresentationBuilder {
    pub fn new(template: SlideTemplate) -> Self {
        Self {
            _template: template,
            slides: Vec::new(),
            metadata: PresentationMetadata {
                title: "Converted Presentation".to_string(),
                author: "md2pptx".to_string(),
                created: Utc::now(),
                modified: Utc::now(),
                slide_count: 0,
            },
        }
    }

    pub fn from_markdown(markdown_doc: &MarkdownDocument, template: SlideTemplate) -> Result<Self> {
        let mut builder = Self::new(template);

        // Set metadata from markdown document
        if let Some(title) = &markdown_doc.metadata.title {
            builder.metadata.title = title.clone();
        }
        if let Some(author) = &markdown_doc.metadata.author {
            builder.metadata.author = author.clone();
        }

        // Convert markdown slides to PPTX slides
        for slide in &markdown_doc.slides {
            builder.add_slide_from_markdown(slide)?;
        }

        builder.metadata.slide_count = builder.slides.len();
        Ok(builder)
    }

    fn add_slide_from_markdown(&mut self, slide: &Slide) -> Result<()> {
        let mut pptx_slide = PptxSlide {
            _id: Uuid::new_v4().to_string(),
            title: slide.title.clone(),
            content: Vec::new(),
        };

        for element in &slide.content {
            match element {
                SlideElement::Heading { level: _, text } => {
                    pptx_slide.content.push(PptxElement::Text(text.clone()));
                }
                SlideElement::Paragraph { text } => {
                    pptx_slide.content.push(PptxElement::Text(text.clone()));
                }
                SlideElement::List { items, ordered } => {
                    if *ordered {
                        pptx_slide
                            .content
                            .push(PptxElement::NumberedList(items.clone()));
                    } else {
                        pptx_slide
                            .content
                            .push(PptxElement::BulletList(items.clone()));
                    }
                }
                SlideElement::CodeBlock { language, code } => {
                    pptx_slide.content.push(PptxElement::Code {
                        _language: language.clone(),
                        content: code.clone(),
                    });
                }
                SlideElement::Image { alt_text, url } => {
                    pptx_slide.content.push(PptxElement::Image {
                        alt: alt_text.clone(),
                        _url: url.clone(),
                    });
                }
                SlideElement::Table { headers, rows } => {
                    pptx_slide.content.push(PptxElement::Table {
                        headers: headers.clone(),
                        _rows: rows.clone(),
                    });
                }
                SlideElement::Quote { text } => {
                    pptx_slide.content.push(PptxElement::Quote(text.clone()));
                }
            }
        }

        self.slides.push(pptx_slide);
        Ok(())
    }

    pub fn build(&self) -> Result<Vec<u8>> {
        let mut buffer = Vec::new();
        {
            let cursor = Cursor::new(&mut buffer);
            let mut zip = ZipWriter::new(cursor);

            // Add required PowerPoint files
            self.add_content_types(&mut zip)?;
            self.add_relationships(&mut zip)?;
            self.add_app_properties(&mut zip)?;
            self.add_core_properties(&mut zip)?;
            self.add_presentation(&mut zip)?;
            self.add_presentation_relationships(&mut zip)?;
            self.add_slide_master(&mut zip)?;
            self.add_slide_master_relationships(&mut zip)?;
            self.add_slide_layout(&mut zip)?;

            // Add slides
            for (index, slide) in self.slides.iter().enumerate() {
                self.add_slide(&mut zip, slide, index + 1)?;
                self.add_slide_relationships(&mut zip, slide, index + 1)?;
            }

            self.add_theme(&mut zip)?;

            zip.finish()?;
        }
        Ok(buffer)
    }

    fn add_content_types(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let content_types = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
    <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
    <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>{}</Types>"#,
            self.slides.iter().enumerate().map(|(i, _)| {
                format!(r#"
    <Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#, i + 1)
            }).collect::<String>()
        );

        zip.start_file("[Content_Types].xml", Default::default())?;
        zip.write_all(content_types.as_bytes())?;
        Ok(())
    }

    fn add_relationships(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;

        zip.start_file("_rels/.rels", Default::default())?;
        zip.write_all(relationships.as_bytes())?;
        Ok(())
    }

    fn add_app_properties(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let app_props = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>md2pptx</Application>
    <PresentationFormat>On-screen Show (4:3)</PresentationFormat>
    <Slides>{}</Slides>
    <Notes>0</Notes>
    <HiddenSlides>0</HiddenSlides>
    <MMClips>0</MMClips>
    <ScaleCrop>false</ScaleCrop>
    <Company>md2pptx</Company>
    <AppVersion>16.0000</AppVersion>
</Properties>"#,
            self.metadata.slide_count
        );

        zip.start_file("docProps/app.xml", Default::default())?;
        zip.write_all(app_props.as_bytes())?;
        Ok(())
    }

    fn add_core_properties(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let core_props = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>{}</dc:title>
    <dc:creator>{}</dc:creator>
    <dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
</cp:coreProperties>"#,
            escape_xml(&self.metadata.title),
            escape_xml(&self.metadata.author),
            self.metadata.created.format("%Y-%m-%dT%H:%M:%SZ"),
            self.metadata.modified.format("%Y-%m-%dT%H:%M:%SZ")
        );

        zip.start_file("docProps/core.xml", Default::default())?;
        zip.write_all(core_props.as_bytes())?;
        Ok(())
    }

    fn add_presentation(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let slide_id_list = self
            .slides
            .iter()
            .enumerate()
            .map(|(i, _slide)| format!(r#"<p:sldId id="{}" r:id="rId{}"/>"#, 256 + i, i + 2))
            .collect::<String>();

        let presentation = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:sldMasterIdLst>
        <p:sldMasterId id="2147483648" r:id="rId1"/>
    </p:sldMasterIdLst>
    <p:sldIdLst>
        {}
    </p:sldIdLst>
    <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
    <p:notesSz cx="6858000" cy="9144000"/>
    <p:defaultTextStyle>
        <a:defPPr>
            <a:defRPr lang="en-US"/>
        </a:defPPr>
    </p:defaultTextStyle>
</p:presentation>"#,
            slide_id_list
        );

        zip.start_file("ppt/presentation.xml", Default::default())?;
        zip.write_all(presentation.as_bytes())?;
        Ok(())
    }

    fn add_presentation_relationships(
        &self,
        zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>,
    ) -> Result<()> {
        let slide_relationships = self.slides.iter().enumerate().map(|(i, _)| {
            format!(r#"    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#, i + 2, i + 1)
        }).collect::<String>();

        let next_id = self.slides.len() + 2;
        let relationships = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
{}    <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#,
            slide_relationships, next_id
        );

        zip.start_file("ppt/_rels/presentation.xml.rels", Default::default())?;
        zip.write_all(relationships.as_bytes())?;
        Ok(())
    }

    fn add_slide_master(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let slide_master = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
        </p:spTree>
    </p:cSld>
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
    </p:sldLayoutIdLst>
    <p:txStyles>
        <p:titleStyle>
            <a:lvl1pPr>
                <a:defRPr sz="4400" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1"/>
                    </a:solidFill>
                    <a:latin typeface="+mj-lt"/>
                    <a:ea typeface="+mj-ea"/>
                    <a:cs typeface="+mj-cs"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:titleStyle>
        <p:bodyStyle>
            <a:lvl1pPr>
                <a:defRPr sz="2800">
                    <a:solidFill>
                        <a:schemeClr val="tx1"/>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:bodyStyle>
        <p:otherStyle>
            <a:lvl1pPr>
                <a:defRPr>
                    <a:solidFill>
                        <a:schemeClr val="tx1"/>
                    </a:solidFill>
                    <a:latin typeface="+mn-lt"/>
                    <a:ea typeface="+mn-ea"/>
                    <a:cs typeface="+mn-cs"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:otherStyle>
    </p:txStyles>
</p:sldMaster>"#;

        zip.start_file("ppt/slideMasters/slideMaster1.xml", Default::default())?;
        zip.write_all(slide_master.as_bytes())?;
        Ok(())
    }

    fn add_slide_master_relationships(
        &self,
        zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>,
    ) -> Result<()> {
        let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;

        zip.start_file(
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            Default::default(),
        )?;
        zip.write_all(relationships.as_bytes())?;
        Ok(())
    }

    fn add_slide_layout(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let slide_layout = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">
    <p:cSld name="Title Slide">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ctrTitle"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Subtitle 2"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="subTitle" idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>Click to edit Master subtitle style</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sldLayout>"#;

        zip.start_file("ppt/slideLayouts/slideLayout1.xml", Default::default())?;
        zip.write_all(slide_layout.as_bytes())?;
        Ok(())
    }

    fn add_slide(
        &self,
        zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>,
        slide: &PptxSlide,
        slide_num: usize,
    ) -> Result<()> {
        let title_text = slide.title.as_deref().unwrap_or("Slide Title");
        let content_shapes = self.generate_content_shapes(&slide.content);

        let slide_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="457200"/>
                        <a:ext cx="7772400" cy="1143000"/>
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>{}</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>
                </p:txBody>
            </p:sp>
            {}
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping/>
    </p:clrMapOvr>
</p:sld>"#,
            escape_xml(title_text),
            content_shapes
        );

        zip.start_file(
            format!("ppt/slides/slide{}.xml", slide_num),
            Default::default(),
        )?;
        zip.write_all(slide_xml.as_bytes())?;
        Ok(())
    }

    fn generate_content_shapes(&self, content: &[PptxElement]) -> String {
        if content.is_empty() {
            return String::new();
        }

        let mut shapes = String::new();
        let mut shape_id = 3;
        let mut y_pos = 1828800; // Starting Y position below title

        for element in content {
            match element {
                PptxElement::Text(text) | PptxElement::Quote(text) => {
                    shapes.push_str(&format!(
                        r#"
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{}" name="Content {}"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="{}"/>
                        <a:ext cx="7772400" cy="1200000"/>
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>{}</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>
                </p:txBody>
            </p:sp>"#,
                        shape_id,
                        shape_id,
                        y_pos,
                        escape_xml(text)
                    ));

                    shape_id += 1;
                    y_pos += 400000;
                }
                PptxElement::BulletList(items) | PptxElement::NumberedList(items) => {
                    let list_items = items
                        .iter()
                        .map(|item| {
                            format!(
                                r#"
                    <a:p>
                        <a:pPr lvl="0"/>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>{}</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>"#,
                                escape_xml(item)
                            )
                        })
                        .collect::<String>();

                    shapes.push_str(&format!(
                        r#"
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{}" name="Content {}"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="{}"/>
                        <a:ext cx="7772400" cy="1600000"/>
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    {}
                </p:txBody>
            </p:sp>"#,
                        shape_id, shape_id, y_pos, list_items
                    ));

                    shape_id += 1;
                    y_pos += 500000;
                }
                PptxElement::Code { content, .. } => {
                    shapes.push_str(&format!(
                        r#"
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{}" name="Code {}"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="2"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="{}"/>
                        <a:ext cx="7772400" cy="1200000"/>
                    </a:xfrm>
                    <a:solidFill>
                        <a:srgbClr val="F8F8F8"/>
                    </a:solidFill>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US">
                                <a:latin typeface="Consolas"/>
                            </a:rPr>
                            <a:t>{}</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                    </a:p>
                </p:txBody>
            </p:sp>"#,
                        shape_id,
                        shape_id,
                        y_pos,
                        escape_xml(content)
                    ));

                    shape_id += 1;
                    y_pos += 600000;
                }
                _ => {
                    // For now, convert other elements to text
                    let text = match element {
                        PptxElement::Image { alt, .. } => format!("[Image: {}]", alt),
                        PptxElement::Table { headers, .. } => {
                            format!("[Table: {}]", headers.join(", "))
                        }
                        _ => "[Unsupported element]".to_string(),
                    };

                    shapes.push_str(&format!(
                        r#"
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{}" name="Content {}"/>
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1"/>
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="2"/>
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="{}"/>
                        <a:ext cx="7772400" cy="600000"/>
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US"/>
                            <a:t>{}</a:t>
                        </a:r>
                        <a:endParaRPr lang="en-US"/>
                        </a:p>
                        </p:txBody>
                        </p:sp>"#,
                        shape_id,
                        shape_id,
                        y_pos,
                        escape_xml(&text)
                    ));

                    shape_id += 1;
                    y_pos += 400000;
                }
            }
        }

        shapes
    }

    fn add_slide_relationships(
        &self,
        zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>,
        _slide: &PptxSlide,
        slide_num: usize,
    ) -> Result<()> {
        let relationships = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

        zip.start_file(
            format!("ppt/slides/_rels/slide{}.xml.rels", slide_num),
            Default::default(),
        )?;
        zip.write_all(relationships.as_bytes())?;
        Ok(())
    }

    fn add_theme(&self, zip: &mut ZipWriter<Cursor<&mut Vec<u8>>>) -> Result<()> {
        let theme = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:sysClr val="windowText" lastClr="000000"/>
            </a:dk1>
            <a:lt1>
                <a:sysClr val="window" lastClr="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="1F497D"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="EEECE1"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="4F81BD"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="F79646"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="9BBB59"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="8064A2"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4BACC6"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="F39646"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0000FF"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="800080"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="35000">
                            <a:schemeClr val="phClr">
                                <a:tint val="37000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="15000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="1"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:shade val="51000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="80000">
                            <a:schemeClr val="phClr">
                                <a:shade val="93000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="94000"/>
                                <a:satMod val="135000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr">
                            <a:shade val="95000"/>
                            <a:satMod val="105000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="38000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="40000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="40000">
                            <a:schemeClr val="phClr">
                                <a:tint val="45000"/>
                                <a:shade val="99000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="20000"/>
                                <a:satMod val="255000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
                    </a:path>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="80000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="30000"/>
                                <a:satMod val="200000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                    </a:path>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
</a:theme>"#;

        zip.start_file("ppt/theme/theme1.xml", Default::default())?;
        zip.write_all(theme.as_bytes())?;
        Ok(())
    }
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::presentation::templates::SlideTemplate;

    #[test]
    fn test_escape_xml() {
        assert_eq!(escape_xml("Hello & <world>"), "Hello &amp; &lt;world&gt;");
        assert_eq!(
            escape_xml("\"quoted\" & 'apostrophe'"),
            "&quot;quoted&quot; &amp; &apos;apostrophe&apos;"
        );
    }

    #[test]
    fn test_presentation_builder_creation() {
        let builder = PresentationBuilder::new(SlideTemplate::Default);
        assert_eq!(builder.slides.len(), 0);
        assert_eq!(builder.metadata.title, "Converted Presentation");
    }
}
