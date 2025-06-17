# Sample Presentation

Welcome to our sample presentation created from Markdown!

## Introduction

This is the introduction slide with some basic content:

- Point 1: Markdown is easy to write
- Point 2: PowerPoint is great for presentations  
- Point 3: md2pptx combines both!

## Code Example

Here's how you can include code in your slides:

```rust
fn main() {
    println!("Hello, md2pptx!");
    
    // Convert markdown to PowerPoint
    let result = convert_markdown_to_pptx("input/", "output.pptx", "default", false);
    match result {
        Ok(_) => println!("Conversion successful!"),
        Err(e) => eprintln!("Error: {}", e),
    }
}
```

## Tables and Data

You can also include tables:

| Feature | Status | Priority |
|---------|--------|----------|
| Markdown parsing | ✅ Complete | High |
| PowerPoint generation | ✅ Complete | High |
| Multiple templates | ✅ Complete | Medium |
| Image support | 🔄 In Progress | Medium |

## Quotes and Emphasis

> "The best way to predict the future is to create it." - Peter Drucker

This slide demonstrates:

- **Bold text** for emphasis
- *Italic text* for subtle emphasis
- `Inline code` for technical terms

## Conclusion

Thank you for using md2pptx! 

### Key Benefits:

1. **Easy authoring** - Write in familiar Markdown syntax
2. **Professional output** - Generate polished PowerPoint presentations
3. **Version control friendly** - Track changes in plain text
4. **Automation ready** - Perfect for CI/CD pipelines

Visit our GitHub repository for more information and examples.
