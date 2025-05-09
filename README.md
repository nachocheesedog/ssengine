# SSEngine: Spreadsheet Engine for AI Agents

[![License: MIT OR Apache-2.0](https://img.shields.io/badge/License-MIT%20OR%20Apache--2.0-blue.svg)](LICENSE-MIT)

A pure Rust spreadsheet engine optimized for AI agent interactions. SSEngine enables AI systems to programmatically build and manipulate spreadsheets, including complex financial models like 3-statement models and DCF analyses.

## Features

- **Full-featured formula engine** compatible with standard spreadsheet functions
- **Multi-sheet support** with cross-referencing
- **AI-friendly JSON API** for building and modifying workbooks
- **XLSX export** for interoperability with Excel and Google Sheets
- **Finance-oriented functions** (NPV, IRR, XIRR, etc.)
- **Built with Rust** for performance, safety, and reliability

## Project Structure

SSEngine is organized as a cargo workspace with these components:

- **`core/`**: Data model, parsing, and evaluation engine
- **`io/`**: XLSX and CSV import/export functionality  
- **`sdk/`**: HTTP/JSON SDK for AI agents
- **`cli/`**: Command-line interface
- **`examples/`**: Sample workbooks and integration tests

## Getting Started

### Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) 1.70+

### Building from Source

```bash
# Clone the repository
git clone https://github.com/user/ssengine.git
cd ssengine

# Build the project
cargo build --release

# Run the CLI
./target/release/ssengine-cli --help
```

### Using the API Server

```bash
# Start the API server on localhost:8080
cargo run --release -p ssengine-cli -- serve
```

### API Examples

Create a workbook with a formula:

```bash
# Create a new sheet
curl -X POST http://localhost:8080/add_sheet -d '{"name":"Sheet1"}'

# Set cell values
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":0,"col":0,"value":"Revenue"}'
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":0,"col":1,"value":"100"}'
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":1,"col":0,"value":"Expenses"}'
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":1,"col":1,"value":"75"}'
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":2,"col":0,"value":"Profit"}'
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":2,"col":1,"value":"=B1-B2"}'

# Export to XLSX
curl -X POST http://localhost:8080/export_xlsx -d '{"path":"simple_model.xlsx"}'
```

## Documentation

- [Design Philosophy](docs.md)
- [API Reference](specs.md)
- [Development Roadmap](todo.md)

## License

SSEngine is dual-licensed under either:

- MIT License ([LICENSE-MIT](LICENSE-MIT) or http://opensource.org/licenses/MIT)
- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE) or http://www.apache.org/licenses/LICENSE-2.0)

at your option.

## Acknowledgments

- This project uses several Rust crates with compatible licenses:
  - `rust_xlsxwriter` for XLSX generation
  - `calamine` for XLSX reading
  - `pest` for formula parsing
  - `petgraph` for the dependency graph
