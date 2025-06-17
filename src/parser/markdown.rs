use crate::utils::error::Error;
use crate::Result;
use pulldown_cmark::{Event, HeadingLevel, Parser, Tag};
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub struct MarkdownDocument {
    pub slides: Vec<Slide>,
    pub metadata: DocumentMetadata,
}

#[derive(Debug, Clone)]
pub struct Slide {
    pub title: Option<String>,
    pub content: Vec<SlideElement>,
}

#[derive(Debug, Clone)]
pub enum SlideElement {
    Heading {
        level: u8,
        text: String,
    },
    Paragraph {
        text: String,
    },
    List {
        items: Vec<String>,
        ordered: bool,
    },
    CodeBlock {
        language: Option<String>,
        code: String,
    },
    Image {
        alt_text: String,
        url: String,
    },
    Table {
        headers: Vec<String>,
        rows: Vec<Vec<String>>,
    },
    Quote {
        text: String,
    },
}

#[derive(Debug, Clone, Default)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub description: Option<String>,
    pub custom_properties: HashMap<String, String>,
}

impl MarkdownDocument {
    pub fn parse(markdown_content: &str) -> Result<Self> {
        let parser = Parser::new(markdown_content);
        let mut document = MarkdownDocument {
            slides: Vec::new(),
            metadata: DocumentMetadata::default(),
        };

        let mut current_slide = Slide {
            title: None,
            content: Vec::new(),
        };

        let events: Vec<Event> = parser.collect();
        let mut i = 0;

        // Extract metadata from front matter if present
        if let Some(title) = extract_title_from_events(&events) {
            document.metadata.title = Some(title);
        }

        while i < events.len() {
            match &events[i] {
                Event::Start(Tag::Heading(level, _, _)) => {
                    let heading_text = extract_text_from_heading(&events, &mut i)?;

                    match *level {
                        HeadingLevel::H1 => {
                            // H1 creates a new slide with title
                            if !current_slide.content.is_empty() || current_slide.title.is_some() {
                                document.slides.push(current_slide);
                                current_slide = Slide {
                                    title: None,
                                    content: Vec::new(),
                                };
                            }
                            current_slide.title = Some(heading_text);
                        }
                        HeadingLevel::H2 => {
                            // H2 also creates a new slide with title
                            if !current_slide.content.is_empty() || current_slide.title.is_some() {
                                document.slides.push(current_slide);
                                current_slide = Slide {
                                    title: None,
                                    content: Vec::new(),
                                };
                            }
                            current_slide.title = Some(heading_text);
                        }
                        level => {
                            // H3+ becomes content within the slide
                            let level_num = match level {
                                HeadingLevel::H1 => 1,
                                HeadingLevel::H2 => 2,
                                HeadingLevel::H3 => 3,
                                HeadingLevel::H4 => 4,
                                HeadingLevel::H5 => 5,
                                HeadingLevel::H6 => 6,
                            };
                            current_slide.content.push(SlideElement::Heading {
                                level: level_num,
                                text: heading_text,
                            });
                        }
                    }
                }
                Event::Start(Tag::Paragraph) => {
                    let paragraph_text = extract_paragraph_text(&events, &mut i)?;
                    if !paragraph_text.trim().is_empty() {
                        current_slide.content.push(SlideElement::Paragraph {
                            text: paragraph_text,
                        });
                    }
                }
                Event::Start(Tag::List(start_num)) => {
                    let (items, _ordered) = extract_list_items(&events, &mut i)?;
                    current_slide.content.push(SlideElement::List {
                        items,
                        ordered: start_num.is_some(),
                    });
                }
                Event::Start(Tag::CodeBlock(kind)) => {
                    let (code, language) = extract_code_block(&events, &mut i, kind.clone())?;
                    current_slide
                        .content
                        .push(SlideElement::CodeBlock { language, code });
                }
                Event::Start(Tag::Image(_, url, alt)) => {
                    current_slide.content.push(SlideElement::Image {
                        alt_text: alt.to_string(),
                        url: url.to_string(),
                    });
                    i += 1; // Skip the image event
                }
                Event::Start(Tag::BlockQuote) => {
                    let quote_text = extract_quote_text(&events, &mut i)?;
                    current_slide
                        .content
                        .push(SlideElement::Quote { text: quote_text });
                }
                Event::Start(Tag::Table(_)) => {
                    let (headers, rows) = extract_table_data(&events, &mut i)?;
                    current_slide
                        .content
                        .push(SlideElement::Table { headers, rows });
                }
                _ => {
                    i += 1;
                }
            }
        }

        // Add the last slide if it has content
        if !current_slide.content.is_empty() || current_slide.title.is_some() {
            document.slides.push(current_slide);
        }

        if document.slides.is_empty() {
            return Err(Error::markdown_parsing(
                "No slides found in markdown document",
            ));
        }

        Ok(document)
    }
}

fn extract_title_from_events(events: &[Event]) -> Option<String> {
    for (i, event) in events.iter().enumerate() {
        if let Event::Start(Tag::Heading(level, _, _)) = event {
            if *level == HeadingLevel::H1 {
                if let Some(Event::Text(text)) = events.get(i + 1) {
                    return Some(text.to_string());
                }
            }
        }
    }
    None
}

fn extract_text_from_heading(events: &[Event], index: &mut usize) -> Result<String> {
    *index += 1; // Skip the Start(Heading) event
    let mut text = String::new();

    while *index < events.len() {
        match &events[*index] {
            Event::Text(t) => text.push_str(t),
            Event::Code(t) => text.push_str(t),
            Event::End(Tag::Heading(_, _, _)) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok(text)
}

fn extract_paragraph_text(events: &[Event], index: &mut usize) -> Result<String> {
    *index += 1; // Skip the Start(Paragraph) event
    let mut text = String::new();

    while *index < events.len() {
        match &events[*index] {
            Event::Text(t) => text.push_str(t),
            Event::Code(t) => {
                text.push('`');
                text.push_str(t);
                text.push('`');
            }
            Event::Start(Tag::Strong) => text.push_str("**"),
            Event::End(Tag::Strong) => text.push_str("**"),
            Event::Start(Tag::Emphasis) => text.push('*'),
            Event::End(Tag::Emphasis) => text.push('*'),
            Event::End(Tag::Paragraph) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok(text)
}

fn extract_list_items(events: &[Event], index: &mut usize) -> Result<(Vec<String>, bool)> {
    *index += 1; // Skip the Start(List) event
    let mut items = Vec::new();
    let mut current_item = String::new();
    let ordered = false;

    while *index < events.len() {
        match &events[*index] {
            Event::Start(Tag::Item) => {
                if !current_item.is_empty() {
                    items.push(current_item.trim().to_string());
                    current_item.clear();
                }
            }
            Event::Text(t) => current_item.push_str(t),
            Event::Code(t) => {
                current_item.push('`');
                current_item.push_str(t);
                current_item.push('`');
            }
            Event::End(Tag::Item) => {
                if !current_item.is_empty() {
                    items.push(current_item.trim().to_string());
                    current_item.clear();
                }
            }
            Event::End(Tag::List(_)) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok((items, ordered))
}

fn extract_code_block(
    events: &[Event],
    index: &mut usize,
    kind: pulldown_cmark::CodeBlockKind,
) -> Result<(String, Option<String>)> {
    let language = match kind {
        pulldown_cmark::CodeBlockKind::Fenced(lang) => {
            if lang.is_empty() {
                None
            } else {
                Some(lang.to_string())
            }
        }
        pulldown_cmark::CodeBlockKind::Indented => None,
    };

    *index += 1; // Skip the Start(CodeBlock) event
    let mut code = String::new();

    while *index < events.len() {
        match &events[*index] {
            Event::Text(t) => code.push_str(t),
            Event::End(Tag::CodeBlock(_)) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok((code, language))
}

fn extract_quote_text(events: &[Event], index: &mut usize) -> Result<String> {
    *index += 1; // Skip the Start(BlockQuote) event
    let mut text = String::new();

    while *index < events.len() {
        match &events[*index] {
            Event::Text(t) => text.push_str(t),
            Event::Code(t) => {
                text.push('`');
                text.push_str(t);
                text.push('`');
            }
            Event::End(Tag::BlockQuote) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok(text)
}

fn extract_table_data(
    events: &[Event],
    index: &mut usize,
) -> Result<(Vec<String>, Vec<Vec<String>>)> {
    *index += 1; // Skip the Start(Table) event
    let mut headers = Vec::new();
    let mut rows = Vec::new();
    let mut current_row = Vec::new();
    let mut current_cell = String::new();
    let mut in_header = true;

    while *index < events.len() {
        match &events[*index] {
            Event::Start(Tag::TableHead) => {
                in_header = true;
            }
            Event::End(Tag::TableHead) => {
                if !current_cell.is_empty() {
                    headers.push(current_cell.trim().to_string());
                    current_cell.clear();
                }
                if !current_row.is_empty() {
                    headers.append(&mut current_row);
                }
                in_header = false;
            }
            Event::Start(Tag::TableRow) => {
                current_row.clear();
            }
            Event::End(Tag::TableRow) => {
                if !current_cell.is_empty() {
                    current_row.push(current_cell.trim().to_string());
                    current_cell.clear();
                }
                if !in_header && !current_row.is_empty() {
                    rows.push(current_row.clone());
                }
            }
            Event::Start(Tag::TableCell) => {
                current_cell.clear();
            }
            Event::End(Tag::TableCell) => {
                if in_header {
                    headers.push(current_cell.trim().to_string());
                } else {
                    current_row.push(current_cell.trim().to_string());
                }
                current_cell.clear();
            }
            Event::Text(t) => current_cell.push_str(t),
            Event::End(Tag::Table(_)) => {
                *index += 1;
                break;
            }
            _ => {}
        }
        *index += 1;
    }

    Ok((headers, rows))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_markdown() {
        let markdown = r#"# Title Slide

This is the first slide content.

## Second Slide

- Item 1
- Item 2
- Item 3

### Subsection

More content here.
"#;

        let doc = MarkdownDocument::parse(markdown).unwrap();
        assert_eq!(doc.slides.len(), 2);
        assert_eq!(doc.slides[0].title, Some("Title Slide".to_string()));
        assert_eq!(doc.slides[1].title, Some("Second Slide".to_string()));
    }

    #[test]
    fn test_parse_code_block() {
        let markdown = r#"# Code Example

```rust
fn main() {
    println!("Hello, world!");
}
```
"#;

        let doc = MarkdownDocument::parse(markdown).unwrap();
        assert_eq!(doc.slides.len(), 1);

        if let SlideElement::CodeBlock { language, code } = &doc.slides[0].content[0] {
            assert_eq!(language, &Some("rust".to_string()));
            assert!(code.contains("fn main()"));
        } else {
            panic!("Expected code block");
        }
    }
}
