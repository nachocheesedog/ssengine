# Spreadsheet Engine – Project TODO

## Table of Contents
1. [High-Level Roadmap](#high-level-roadmap)
2. [Detailed Task Backlog](#detailed-task-backlog)
   1. [Core Engine](#core-engine)
   2. [I/O – XLSX + CSV](#io)
   3. [Formula Library](#formula-library)
   4. [AI Tooling / SDK](#ai-tooling)
   5. [Testing & QA](#testing--qa)
   6. [Dev Experience](#dev-experience)
   7. [Performance](#performance)
   8. [Security & Compliance](#security--compliance)
   9. [Documentation](#documentation)
3. [Design References](#design-references)
4. [Security Checklist](#security-checklist)
5. [Design Decisions & Architecture](#design-decisions--architecture)
6. [Execution Plan & Milestones](#execution-plan--milestones)

---

## High-Level Roadmap *(v2025-05-08)*
| Priority | Epic | Outcome |
|---|---|---|
| P0 | Minimal Viable Engine | Parse & evaluate formulas, write multi-sheet XLSX |
| P0 | AI SDK | JSON-RPC interface callable by LLM-agents |
| P1 | Finance Function Pack | NPV, IRR, PMT, PV, FV implemented. Examples in dcf_model.rs |
| P1 | Import Support | Read XLSX/CSV for context enrichment |
| P2 | Optimisations | Incremental recalc, parallelism |
| P2 | Expanded Function Pack | Statistics, text, lookup |
| P3 | Streaming Mode | Headless server that keeps workbook in memory |

---

## Detailed Task Backlog

> **Notation** – Each task carries an *ID*, creation **Date**, and _Status_.
> Status values: TODO / WIP / BLOCKED / DONE / DEFER.

### Core Engine <a id="core-engine"></a>
| ID | Date | Status | Task |
|---|---|---|---|
| CE-01 | 2025-05-08 | TODO | Choose workbook data model (sparse vs dense, columnar vs row-major). |
| CE-02 | 2025-05-08 | TODO | Define `Cell`, `Sheet`, `Workbook` structs – follow SRP. |
| CE-03 | 2025-05-08 | TODO | Implement dependency graph for formula recalculation. |
| CE-04 | 2025-05-08 | TODO | Implement formula parser (PEG / pest) yielding AST. |
| CE-05 | 2025-05-08 | DONE | Implement evaluator w/ basic math & references. |
| CE-06 | 2025-05-08 | DONE | Handle cross-sheet references. |
| CE-07 | 2025-05-08 | DONE | Error propagation (DIV/0!, #REF!, etc.). |
| CE-08 | 2025-05-08 | DONE | Circular reference detection. |

### I/O – XLSX + CSV <a id="io"></a>
| IO-01 | 2025-05-08 | TODO | Add XLSX writer using `rust_xlsxwriter`. |
| IO-02 | 2025-05-08 | TODO | Support workbook styles minimal subset. |
| IO-03 | 2025-05-08 | TODO | Add XLSX reader via `calamine` (MIT). |
| IO-04 | 2025-05-08 | TODO | CSV import/export helpers. |

### Formula Library <a id="formula-library"></a>
| FL-01 | 2025-05-08 | DONE | Implement arithmetic + `SUM`, `AVERAGE`, `MIN`, `MAX`. |
| FL-02 | 2025-05-08 | DONE | Implement logical functions (`IF`, `AND`, `OR`, `NOT`). |
| FL-03 | 2025-05-08 | DONE | Lookup functions (`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`). |
| FL-04 | 2025-05-08 | DONE | Date/time functions (`,YEAR`, `MONTH`, `DATEDIF`). |
| FL-05 | 2025-05-08 | DONE | Finance pack: `NPV`, `IRR`, `XNPV`, `XIRR`. |
| FL-06 | 2025-05-08 | DONE | Text functions (`LEFT`, `RIGHT`, `CONCAT`). |
| FL-07 | 2025-05-08 | DONE | Conditional aggregates (`SUMIF`, `COUNTIF`, `AVERAGEIF`, `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`). |
| FL-08 | 2025-05-08 | DONE | Error handling functions (`IFERROR`, `IFNA`, `IFS`). |
| FL-09 | 2025-05-08 | DONE | Advanced lookup functions (`XLOOKUP`, `XMATCH`, `OFFSET`, `INDIRECT`). |
| FL-10 | 2025-05-08 | DONE | Advanced date functions (`EOMONTH`, `EDATE`, `NETWORKDAYS`, `WORKDAY`, `YEARFRAC`). |
| FL-11 | 2025-05-08 | DONE | Math functions (`MOD`, `CEILING`, `FLOOR`, `MROUND`, `LOG`, `LN`, `EXP`). |
| FL-12 | 2025-05-08 | DONE | Random functions (`RAND`, `RANDBETWEEN`, `RANDARRAY`). |
| FL-13 | 2025-05-08 | DONE | Statistical functions (`MODE.SNGL`, `COVARIANCE.P`, `CORREL`, `AGGREGATE`). |
| FL-14 | 2025-05-08 | DONE | Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`, `LAMBDA`). |

### AI Tooling / SDK <a id="ai-tooling"></a>
| AT-01 | 2025-05-08 | TODO | Define JSON schema for tool operations (add_sheet, set_cell, eval_formula, export_xlsx). |
| AT-02 | 2025-05-08 | TODO | Provide thin Rust HTTP (Axum) server exposing endpoints. |
| AT-03 | 2025-05-08 | TODO | Add OpenAI function-calling manifest example. |
| AT-04 | 2025-05-08 | TODO | Autogenerate TypeScript client. |

### Testing & QA <a id="testing--qa"></a>
| TST-01 | 2025-05-08 | TODO | Unit tests for parser, evaluator (>= 90 % coverage target). |
| TST-02 | 2025-05-08 | TODO | Property tests with `proptest` for numeric accuracy. |
| TST-03 | 2025-05-08 | TODO | Golden-file tests: compare output XLSX vs expected. |
| TST-04 | 2025-05-08 | TODO | Benchmarks using `criterion`. |

### Dev Experience <a id="dev-experience"></a>
| DX-01 | 2025-05-08 | TODO | Cargo workspace layout `/core /io /sdk /cli`. |
| DX-02 | 2025-05-08 | TODO | Add `justfile` with common tasks (lint, test, bench). |
| DX-03 | 2025-05-08 | TODO | Integrate `cargo clippy`, `cargo fmt` CI. |
| DX-04 | 2025-05-08 | TODO | GitHub Actions pipeline. |

### Performance <a id="performance"></a>
| PF-01 | 2025-05-08 | TODO | Implement incremental recalculation using topological sort. |
| PF-02 | 2025-05-08 | TODO | Parallel evaluation with Rayon for independent ranges. |
| PF-03 | 2025-05-08 | TODO | Memory profiling & optimisation pass. |

### Security & Compliance <a id="security--compliance"></a>
| SC-01 | 2025-05-08 | TODO | No `unsafe` Rust blocks unless justified & reviewed. |
| SC-02 | 2025-05-08 | TODO | Dependabot audit workflow. |
| SC-03 | 2025-05-08 | TODO | Validate and sanitise all file paths. |
| SC-04 | 2025-05-08 | TODO | Document threat model in `security.md`. |

### Documentation <a id="documentation"></a>
| DOC-01 | 2025-05-08 | TODO | `README.md` – overview, usage, badges. |
| DOC-02 | 2025-05-08 | TODO | `design.md` – architecture diagrams & data flow. |
| DOC-03 | 2025-05-08 | TODO | `specs.md` – formula grammar, API spec. |
| DOC-04 | 2025-05-08 | TODO | Rustdoc examples. |

---

## Design References <a id="design-references"></a>
* Formula grammar – see Excel OpenXML ECMA-376, Part 1, §18.17.
* `rust_xlsxwriter` for writing XLSX – MIT License.
* `calamine` crate for reading XLSX/ODS/CSV – MIT License.
* Similar open-source projects: lotus-spreadsheet, oxcel.

---

## Security Checklist <a id="security-checklist"></a>
- [ ] All crates are MIT/Apache-2.0 compatible.  
- [ ] No secrets committed (use `.env`).  
- [ ] Input parsing fuzz-tested (`cargo-fuzz`).  
- [ ] CI runs `cargo audit` & `cargo deny`.  
- [ ] Exported files sanitized against path traversal.  
- [ ] Web SDK endpoints rate-limited & CORS-controlled.  
- [ ] User-provided formulas executed in-process only – no external eval.  

---

## Design Decisions & Architecture

### Data Model
- **Sparse storage**: Each `Sheet` maintains a `HashMap<(u32, u32), Cell>` to save memory on large but lightly-populated models.
- **Column meta cache**: Separate vector stores column widths, formats, and statistics for fast lookup operations.
- **Cell**: enum‐based value holder (`Number(f64)`, `Text(String)`, `Bool(bool)`, `Datetime(DateTime<Utc>)`, `Error(CellError)`, `Formula(String)`).
- **Workbook**: owns sheets and a global string table to de-duplicate repeated text (similar to XLSX sharedStrings).

### Formula Parsing & Evaluation
- **Grammar**: Excel-compatible grammar expressed in a `pest` file (`excel.pest`).
- **AST**: Simplified node kinds (`Literal`, `Reference`, `UnaryOp`, `BinaryOp`, `FunctionCall`, `Range`).
- **Dependency Graph**: Directed acyclic graph (DAG) keyed by cell coordinates, produced during parse, stored via `petgraph`.
- **Recalc Strategy**: Topological sort with incremental dirty-flag evaluation; fallback full recompute when graph cycles detected.

### Crate Workspace Layout
```
ssengine/
 ├─ core/        # data model, parser, evaluator
 ├─ io/          # xlsx/csv read-write
 ├─ sdk/         # HTTP/JSON SDK for AI agents
 ├─ cli/         # optional CLI utility
 └─ examples/    # sample workbooks & integration tests
```

### Third-Party Libraries (OSI-approved)
| Purpose | Crate | License |
|---|---|---|
| Formula parsing | `pest` | MIT/Apache-2.0 |
| XLSX writing | `rust_xlsxwriter` | MIT |
| XLSX reading | `calamine` | MIT |
| HashMap | `hashbrown` | Apache-2.0 |
| Graph | `petgraph` | MIT |
| HTTP server | `axum` | MIT |
| Async runtime | `tokio` | MIT |
| Testing | `criterion`, `proptest` | Apache-2.0 |

### Concurrency & Performance
- Rayon parallel iterator for evaluating independent subgraphs.
- Interior mutability via `RwLock` at sheet level; read-heavy workloads scale.
- Cache numeric function results for idempotent pure functions (e.g., `NPV`).

### Error Handling
- `thiserror`-based rich error enum; every error convertible to HTTP JSON error for SDK.

---

## Execution Plan & Milestones
| Sprint | Dates | Focus | Key Deliverables |
|---|---|---|---|
| 0 | May 08 → May 14 | Project scaffolding & infrastructure | Cargo workspace, CI pipeline, docs skeleton |
| 1 | May 15 → May 28 | Core engine v0 | Data model, parser v0, evaluator basic math, unit tests |
| 2 | May 29 → Jun 11 | I/O layer | XLSX writer (rust_xlsxwriter), CSV helpers, basic styles |
| 3 | Jun 12 → Jun 25 | Formula library phase 1 | Arithmetic agg functions, logical, lookup v1 |
| 4 | Jun 26 → Jul 09 | AI SDK v1 | JSON schema, Axum server, OpenAI manifest example |
| 5 | Jul 10 → Jul 23 | Finance pack & benchmarks | **In Progress**: NPV/IRR implemented, PV/FV/PMT added, dcf_model.rs example created, remaining items: XIRR, criterion benches, optimisation pass |
| 6 | Jul 24 → Aug 06 | Import support & UX polish | XLSX reader, CLI utility, docs update |
| 7 | Aug 07 → Aug 20 | Parallelism & streaming mode | Incremental recalc, Rayon, in-memory daemon |
| 8 | Aug 21 → Sep 03 | Hardening & release v1.0 | Fuzzing, security audit, full docs, crates.io publish |
