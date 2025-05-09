// ssengine-sdk/src/schemas.rs
// Schema definitions for API requests and responses

use serde::{Serialize, Deserialize};
use serde_json::Value;

#[derive(Debug, Serialize, Deserialize)]
pub struct AddSheetRequest {
    pub name: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct AddSheetResponse {
    pub sheet_id: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SetCellRequest {
    pub sheet: String,
    pub row: u32,
    pub col: u32,
    pub value: String, // Can be formula (starts with '=') or raw value
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SetCellResponse {
    pub success: bool,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct GetCellRequest {
    pub sheet: String,
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct GetCellResponse {
    pub value: Value,       // JSON value representation
    pub formula: Option<String>,
    pub formatted: String,  // Formatted string representation
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExportXlsxRequest {
    pub path: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExportXlsxResponse {
    pub success: bool,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ImportXlsxRequest {
    pub path: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ImportXlsxResponse {
    pub success: bool,
}

// OpenAPI schema generation function
pub fn generate_openapi_schema() -> schemars::schema::RootSchema {
    let mut gen = schemars::gen::SchemaGenerator::default();
    let schema = gen.into_root_schema_for::<AddSheetRequest>();
    schema
}

// Generate function-calling schema for LLM AI agents
pub fn generate_function_schema() -> serde_json::Value {
    serde_json::json!({
        "name": "ssengine",
        "description": "Create and manipulate spreadsheets with the ssengine API",
        "parameters": {
            "type": "object",
            "properties": {
                "operation": {
                    "type": "string",
                    "enum": ["add_sheet", "set_cell", "get_cell", "export_xlsx", "import_xlsx"],
                    "description": "The operation to perform on the spreadsheet engine"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "The name of the sheet to operate on"
                },
                "row": {
                    "type": "integer",
                    "description": "The row index (0-based)"
                },
                "col": {
                    "type": "integer",
                    "description": "The column index (0-based)"
                },
                "value": {
                    "type": "string",
                    "description": "The value to set in the cell. Can be a formula (starts with '=') or raw value"
                },
                "path": {
                    "type": "string",
                    "description": "The file path for import/export operations"
                }
            },
            "required": ["operation"]
        }
    })
}
