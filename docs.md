# Spreadsheet Engine – Technical Documentation

_Last updated: 2025-05-08_

## Table of Contents
1. [Introduction](#introduction)
2. [Data Structures](#data-structures)
3. [Formula Grammar](#formula-grammar)
4. [Evaluation Engine](#evaluation-engine)
5. [I/O Layer](#io-layer)
6. [AI SDK & Tool Schema](#ai-sdk--tool-schema)
7. [Concurrency Model](#concurrency-model)
8. [Error Handling](#error-handling)
9. [Extending the Engine](#extending-the-engine)
10. [References](#references)

---

## Introduction
`ssengine` is a pure-Rust, headless spreadsheet engine built for AI-driven workflows. It lets large-language-model agents generate and manipulate workbooks programmatically and export them to standard `.xlsx` files.

Key goals:
* **Determinism** – Functional evaluation model, no hidden state.
* **Performance** – Incremental recalculation & parallelism.
* **Embeddability** – Small dependency footprint, works in Wasm (future).
* **Open-source** – MIT/Apache-2.0 friendly crates only.

---

## Data Structures
### Workbook
```rust
pub struct Workbook {
    sheets: IndexMap<String, Sheet>,
    shared_strings: StringInterner,
}
```
* `sheets` ordered for stable index references.
* `shared_strings` reduces memory for repeated labels.

### Sheet
```rust
pub struct Sheet {
    name: String,
    cells: HashMap<(u32, u32), Cell>,
    column_meta: Vec<ColumnMeta>,
}
```

### Cell
```rust
pub enum Cell {
    Blank,
    Number(f64),
    Text(Symbol),            // interned string
    Bool(bool),
    Datetime(DateTime<Utc>),
    Error(CellError),
    Formula(FormulaNode),
}
```

---

## Formula Grammar
Defined in [`core/src/grammar/excel.pest`]. Key excerpts:
```pest
expression  = _{ logical_or }
logical_or  = { logical_and ~ ("+" ~ logical_and)* }
reference   = @{ sheet_name? ~ cell_addr ~ range? }
function    = { ident ~ "(" ~ arg_list? ~ ")" }
```
Parsing produces an AST (`FormulaNode`). See `design.md` for full grammar.

---

## Evaluation Engine
1. **Parsing** – Formula text → AST.
2. **Graph Build** – For each formula cell, emit edges to precedents.
3. **Dirty Flagging** – On mutation, mark dependents dirty via DFS.
4. **Recalc** – Topological order execution. Parallel when independencies exist.

### Function Dispatch
Implemented via a `HashMap<&'static str, fn(&[Value]) -> Result<Value>>` inside `core::functions`.
* Pure functions cached by `(name, args)` when deterministic.

---

## I/O Layer
* **Writing** – `rust_xlsxwriter` maps our data model to XLSX parts, streaming rows to keep memory low.
* **Reading** – `calamine` converts external workbooks into our internal model, best-effort mapping of functions (unsupported formulas flagged `#N/A`).

---

## AI SDK & Tool Schema
### JSON Schema (excerpt)
```jsonc
{
  "methods": {
    "add_sheet": {
      "params": { "name": "string" },
      "returns": { "sheet_id": "string" }
    },
    "set_cell": {
      "params": {
        "sheet": "string",
        "row": "integer",
        "col": "integer",
        "value": "string"
      }
    },
    "export_xlsx": {
      "params": { "path": "string" }
    }
  }
}
```
* Served via `axum` JSON endpoints. Example cURL:
```bash
curl -X POST http://localhost:8080/set_cell -d '{"sheet":"Sheet1","row":0,"col":0,"value":"=SUM(1,2)"}'
```

---

## Concurrency Model
* **Sheet-level RwLock** – multiple readers, single writer per sheet.
* **Recalc Pool** – Rayon thread-pool splits evaluation by independent subgraphs.
* **Async HTTP** – `tokio` runtime, non-blocking I/O.

---

## Error Handling
Central `Error` enum converts to:
* Rust `Result` types internally.
* HTTP `application/json` `{ code, message }` externally.
* Excel cell errors (`#DIV/0!`) for workbook representation.

---

## Extending the Engine
Add a new function:
1. Implement `core::functions::my_func`.
2. Add to dispatcher map.
3. Update docs + tests (`tests/function_my_func.rs`).
4. Run `cargo test && cargo fmt`.

---

## References
* ECMA-376 OpenXML spec.
* Excel 2016 function reference.
* Crates: `rust_xlsxwriter`, `calamine`, `pest`, `petgraph`, `axum`, `tokio`.
