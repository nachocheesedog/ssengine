// ssengine-sdk/src/api.rs
// API implementation for workbook operations

use ssengine_core::{Workbook, Sheet, Cell, CellValue, EngineError, CellAddress};
use ssengine_io::{read_xlsx, write_xlsx};
use serde::{Serialize, Deserialize};
use std::path::PathBuf;
use std::sync::{Arc, RwLock};

// Workbook API - thread-safe wrapper around a workbook
pub struct WorkbookApi {
    workbook: Arc<RwLock<Workbook>>,
}

impl WorkbookApi {
    pub fn new() -> Self {
        WorkbookApi {
            workbook: Arc::new(RwLock::new(Workbook::new())),
        }
    }
    
    pub fn from_workbook(workbook: Workbook) -> Self {
        WorkbookApi {
            workbook: Arc::new(RwLock::new(workbook)),
        }
    }
    
    // Create a new sheet
    pub fn add_sheet(&self, name: String) -> Result<String, ApiError> {
        let mut wb = self.workbook.write().map_err(|_| ApiError::LockError)?
        wb.add_sheet(name).map_err(ApiError::EngineError)
    }
    
    // Set a cell value
    pub fn set_cell(&self, sheet: String, row: u32, col: u32, value: String) -> Result<(), ApiError> {
        // To be implemented
        Err(ApiError::NotImplemented("set_cell".into()))
    }
    
    // Get a cell value
    pub fn get_cell(&self, sheet: String, row: u32, col: u32) -> Result<CellResponse, ApiError> {
        // To be implemented
        Err(ApiError::NotImplemented("get_cell".into()))
    }
    
    // Export the workbook to XLSX
    pub fn export_xlsx(&self, path: PathBuf) -> Result<(), ApiError> {
        let wb = self.workbook.read().map_err(|_| ApiError::LockError)?;
        write_xlsx(&wb, path).map_err(ApiError::EngineError)
    }
    
    // Import a workbook from XLSX
    pub fn import_xlsx(&self, path: PathBuf) -> Result<(), ApiError> {
        let imported = read_xlsx(path).map_err(ApiError::EngineError)?;
        let mut wb = self.workbook.write().map_err(|_| ApiError::LockError)?;
        *wb = imported;
        Ok(())
    }
}

// API error type
#[derive(Debug, thiserror::Error)]
pub enum ApiError {
    #[error("Engine error: {0}")]
    EngineError(#[from] EngineError),
    
    #[error("Failed to acquire lock on workbook")]
    LockError,
    
    #[error("Feature not implemented: {0}")]
    NotImplemented(String),
    
    #[error("Invalid request: {0}")]
    InvalidRequest(String),
}

// Response types
#[derive(Debug, Serialize, Deserialize)]
pub struct CellResponse {
    pub value: serde_json::Value,
    pub formula: Option<String>,
    pub formatted: String,
}

// Convert ApiError to HTTP response
impl From<ApiError> for axum::http::StatusCode {
    fn from(err: ApiError) -> Self {
        match err {
            ApiError::EngineError(_) => axum::http::StatusCode::INTERNAL_SERVER_ERROR,
            ApiError::LockError => axum::http::StatusCode::SERVICE_UNAVAILABLE,
            ApiError::NotImplemented(_) => axum::http::StatusCode::NOT_IMPLEMENTED,
            ApiError::InvalidRequest(_) => axum::http::StatusCode::BAD_REQUEST,
        }
    }
}

// Helper function to convert ApiError to JSON response
pub fn error_to_json(err: ApiError) -> serde_json::Value {
    serde_json::json!({
        "error": err.to_string(),
        "code": match err {
            ApiError::EngineError(_) => "ENGINE_ERROR",
            ApiError::LockError => "LOCK_ERROR",
            ApiError::NotImplemented(_) => "NOT_IMPLEMENTED",
            ApiError::InvalidRequest(_) => "INVALID_REQUEST",
        }
    })
}
