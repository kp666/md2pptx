# MD2PPTX Development Rules

## Project Overview
This is a Rust CLI tool that converts Markdown (.md) files from a folder into PowerPoint (.pptx) presentations.

## Development Commands
- **Build**: `cargo build`
- **Test**: `cargo test`
- **Run**: `cargo run -- <input_folder> <output_file>`
- **Lint**: `cargo clippy`
- **Format**: `cargo fmt`
- **Check**: `cargo check`

## Code Organization

### 1. Modular Component Design
Structure code with clear module boundaries:
- Separate the markdown parser from the presentation generator
- Use Rust modules to organize related functionality
- Define clear APIs between components
- Keep CLI interface separate from core logic

### 2. Core Modules Structure
```
src/
├── main.rs           # CLI entry point and argument parsing
├── lib.rs            # Library exports and module declarations
├── parser/           # Markdown parsing functionality
│   ├── mod.rs        # Parser module exports
│   └── markdown.rs   # Markdown AST and parsing logic
├── presentation/     # PowerPoint generation
│   ├── mod.rs        # Presentation module exports
│   ├── builder.rs    # PPTX builder and formatting
│   └── templates.rs  # Slide templates and themes
├── converter/        # Conversion logic
│   ├── mod.rs        # Converter module exports
│   └── md_to_pptx.rs # Main conversion pipeline
└── utils/            # Utility functions
    ├── mod.rs        # Utilities module exports
    └── file_io.rs    # File handling utilities
```

### 3. Error Handling
- Use custom error types with `thiserror` crate
- Implement proper error propagation with `?` operator
- Provide meaningful error messages for CLI users
- Handle file I/O errors gracefully

### 4. CLI Design
- Use `clap` for argument parsing
- Support batch processing of multiple files
- Provide progress indicators for large operations
- Include verbose/quiet output options

## Testing Strategy

### 5. Comprehensive Testing
Implement testing at multiple levels:
- **Unit tests**: Test individual components in isolation
- **Integration tests**: Test the full conversion pipeline
- **CLI tests**: Test command-line interface behavior
- **Regression tests**: Prevent previously fixed bugs
- **Property-based tests**: Use `proptest` for edge cases

### 6. Test Organization
```
tests/
├── integration/      # Integration tests
├── fixtures/         # Test data files
│   ├── input/        # Sample markdown files
│   └── expected/     # Expected output files
└── common/           # Shared test utilities
```

## Documentation

### 7. Code Documentation
Document code thoroughly:
- Explain the "why" behind complex algorithms
- Document invariants and assumptions
- Include examples in public API documentation
- Use `cargo doc` to generate documentation

### 8. User Documentation
- Create comprehensive README.md
- Include usage examples and common patterns
- Document supported Markdown features
- Provide troubleshooting guide

## Rust-Specific Practices

### 9. Memory Safety
- Minimize unsafe code to the smallest possible scope
- Use unsafe blocks only when absolutely necessary
- Validate all preconditions before unsafe operations
- Document safety requirements extensively

### 10. Performance Optimization
Be mindful of Rust-specific performance patterns:
- Avoid unnecessary allocations and cloning
- Use references instead of owned values where appropriate
- Consider using `Cow` for potentially borrowed data
- Profile with `cargo bench` for critical paths
- Use `Arc` and `Rc` judiciously for shared data

### 11. Dependency Management
- Keep dependencies minimal and well-maintained
- Use feature flags to make dependencies optional
- Prefer crates with good documentation and test coverage
- Consider compilation time impact of dependencies

## Implementation Workflow

### 12. Implementation Status Tracking
As the project is developed:
- Study `IMPLEMENTATION_STATUS.md` to determine next steps
- Update `IMPLEMENTATION_STATUS.md` when a step is complete
- Use `IMPLEMENTATION_STATUS.md` to track progress and blockers
- Create GitHub issues for major features and bugs

### 13. Development Workflow
1. **Planning**: Break down features into small, testable units
2. **Implementation**: Write tests first, then implementation
3. **Testing**: Run full test suite before committing
4. **Documentation**: Update docs with new features
5. **Integration**: Test CLI interface with real files

### 14. Code Quality Standards
- Run `cargo clippy` and fix all warnings
- Use `cargo fmt` to maintain consistent formatting
- Maintain test coverage above 80%
- Document all public APIs
- Follow Rust naming conventions

## Project-Specific Rules

### 15. Markdown Processing
- Support CommonMark specification
- Handle various heading levels (H1-H6)
- Process lists, tables, and code blocks
- Support images and links appropriately
- Handle edge cases gracefully

### 16. PowerPoint Generation
- Create clean, professional slide layouts
- Support multiple slide templates
- Handle text formatting (bold, italic, code)
- Ensure proper slide transitions
- Optimize for readability

### 17. File Handling
- Support recursive directory processing
- Handle various file encodings
- Validate input files before processing
- Provide clear error messages for invalid files
- Support both single file and batch processing

### 18. Performance Requirements
- Process large directories efficiently
- Stream large files to avoid memory issues
- Provide progress feedback for long operations
- Support parallel processing where beneficial

## Quality Assurance

### 19. Pre-commit Checks
Before committing code:
- Run `cargo test` (all tests pass)
- Run `cargo clippy` (no warnings)
- Run `cargo fmt --check` (code is formatted)
- Update documentation if needed
- Test CLI interface manually

### 20. Release Criteria
Before releasing:
- All tests pass in CI/CD
- Documentation is up to date
- Performance benchmarks are acceptable
- Security audit passes
- Cross-platform compatibility verified
