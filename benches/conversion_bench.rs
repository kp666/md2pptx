use criterion::{black_box, criterion_group, criterion_main, Criterion};
use md2pptx::{convert_markdown_to_pptx, MarkdownDocument};
use std::fs;
use std::path::Path;
use tempfile::tempdir;

fn benchmark_markdown_parsing(c: &mut Criterion) {
    let sample_markdown = r#"# Performance Test Presentation

This is a performance test slide with various content types.

## Introduction

Welcome to our performance benchmarking presentation.

- Point 1: Basic text processing
- Point 2: List handling
- Point 3: Code block processing

## Code Example

```rust
fn main() {
    println!("Hello, benchmark!");
    for i in 0..100 {
        println!("Processing item {}", i);
    }
}
```

## Data Processing

Here's a table with sample data:

| Item | Value | Status |
|------|-------|--------|
| A    | 100   | Active |
| B    | 200   | Pending|
| C    | 300   | Complete|

## Conclusion

> "Performance matters when dealing with large presentations."

This concludes our benchmark test.
"#;

    c.bench_function("parse_markdown", |b| {
        b.iter(|| MarkdownDocument::parse(black_box(sample_markdown)))
    });
}

fn benchmark_small_presentation_conversion(c: &mut Criterion) {
    let temp_dir = tempdir().unwrap();
    let input_file = temp_dir.path().join("test.md");
    let output_file = temp_dir.path().join("test.pptx");

    let sample_markdown = r#"# Test Presentation

## Slide 1
Content for slide 1

## Slide 2  
Content for slide 2

## Slide 3
Content for slide 3
"#;

    fs::write(&input_file, sample_markdown).unwrap();

    c.bench_function("convert_small_presentation", |b| {
        b.iter(|| {
            // Clean up previous output
            if output_file.exists() {
                fs::remove_file(&output_file).unwrap();
            }

            md2pptx::convert_single_markdown_file(
                black_box(&input_file),
                black_box(&output_file),
                black_box("default"),
                black_box(md2pptx::LogLevel::Quiet),
            )
        })
    });
}

criterion_group!(
    benches,
    benchmark_markdown_parsing,
    benchmark_small_presentation_conversion
);
criterion_main!(benches);
