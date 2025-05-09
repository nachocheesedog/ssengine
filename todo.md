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

---

## High-Level Roadmap *(v2025-05-08)*
| Priority | Epic | Outcome |
|---|---|---|
| P0 | Minimal Viable Engine | Parse & evaluate formulas, write multi-sheet XLSX |
| P0 | AI SDK | JSON-RPC interface callable by LLM-agents |
| P1 | Finance Function Pack | NPV, IRR, XNPV, XIRR, etc. |
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
| CE-05 | 2025-05-08 | TODO | Implement evaluator w/ basic math & references. |
| CE-06 | 2025-05-08 | TODO | Handle cross-sheet references. |
| CE-07 | 2025-05-08 | TODO | Error propagation (DIV/0!, #REF!, etc.). |
| CE-08 | 2025-05-08 | TODO | Circular reference detection. |

### I/O – XLSX + CSV <a id="io"></a>
| IO-01 | 2025-05-08 | TODO | Add XLSX writer using `rust_xlsxwriter`. |
| IO-02 | 2025-05-08 | TODO | Support workbook styles minimal subset. |
| IO-03 | 2025-05-08 | TODO | Add XLSX reader via `calamine` (MIT). |
| IO-04 | 2025-05-08 | TODO | CSV import/export helpers. |

### Formula Library <a id="formula-library"></a>
| FL-01 | 2025-05-08 | TODO | Implement arithmetic + `SUM`, `AVERAGE`, `MIN`, `MAX`. |
| FL-02 | 2025-05-08 | TODO | Implement logical functions (`IF`, `AND`, `OR`, `NOT`). |
| FL-03 | 2025-05-08 | TODO | Lookup functions (`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`). |
| FL-04 | 2025-05-08 | TODO | Date/time functions (`,YEAR`, `MONTH`, `DATEDIF`). |
| FL-05 | 2025-05-08 | TODO | Finance pack: `NPV`, `IRR`, `XNPV`, `XIRR`. |
| FL-06 | 2025-05-08 | TODO | Text functions (`LEFT`, `RIGHT`, `CONCAT`). |

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
