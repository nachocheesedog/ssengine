// ssengine-sdk/src/server.rs
// HTTP server implementation for the spreadsheet API

use crate::api::{WorkbookApi, ApiError, error_to_json};
use crate::schemas::*;
use axum::routing::{get, post};
use axum::{Json, Router, Extension};
use axum::http::StatusCode;
use axum::extract::Path;
use axum::response::{IntoResponse, Response};
use serde_json::{json, Value};
use std::sync::Arc;
use std::net::SocketAddr;
use tower_http::cors::{Any, CorsLayer};

// Initialize the API router
pub fn create_router(api: WorkbookApi) -> Router {
    let api = Arc::new(api);
    
    // Define CORS policy
    let cors = CorsLayer::new()
        .allow_origin(Any)
        .allow_methods(Any)
        .allow_headers(Any);
    
    Router::new()
        // Healthcheck endpoint
        .route("/health", get(health_check))
        
        // Workbook operations
        .route("/add_sheet", post(add_sheet))
        .route("/set_cell", post(set_cell))
        .route("/get_cell", post(get_cell))
        .route("/export_xlsx", post(export_xlsx))
        .route("/import_xlsx", post(import_xlsx))
        
        // Attach shared state and middleware
        .layer(Extension(api))
        .layer(cors)
}

// Run the HTTP server
pub async fn run_server(api: WorkbookApi, addr: SocketAddr) {
    let app = create_router(api);
    
    println!("Starting ssengine SDK server on {}", addr);
    
    axum::Server::bind(&addr)
        .serve(app.into_make_service())
        .await
        .unwrap();
}

// Route handlers

async fn health_check() -> impl IntoResponse {
    Json(json!({ "status": "ok" }))
}

async fn add_sheet(
    Extension(api): Extension<Arc<WorkbookApi>>,
    Json(payload): Json<AddSheetRequest>,
) -> Result<Json<AddSheetResponse>, ApiErrorResponse> {
    let sheet_id = api.add_sheet(payload.name)
        .map_err(ApiErrorResponse)?
    
    Ok(Json(AddSheetResponse { sheet_id }))
}

async fn set_cell(
    Extension(api): Extension<Arc<WorkbookApi>>,
    Json(payload): Json<SetCellRequest>,
) -> Result<Json<SetCellResponse>, ApiErrorResponse> {
    api.set_cell(payload.sheet, payload.row, payload.col, payload.value)
        .map_err(ApiErrorResponse)?;
    
    Ok(Json(SetCellResponse { success: true }))
}

async fn get_cell(
    Extension(api): Extension<Arc<WorkbookApi>>,
    Json(payload): Json<GetCellRequest>,
) -> Result<Json<GetCellResponse>, ApiErrorResponse> {
    let cell = api.get_cell(payload.sheet, payload.row, payload.col)
        .map_err(ApiErrorResponse)?;
    
    Ok(Json(GetCellResponse { 
        value: cell.value,
        formula: cell.formula,
        formatted: cell.formatted,
    }))
}

async fn export_xlsx(
    Extension(api): Extension<Arc<WorkbookApi>>,
    Json(payload): Json<ExportXlsxRequest>,
) -> Result<Json<ExportXlsxResponse>, ApiErrorResponse> {
    api.export_xlsx(payload.path.into())
        .map_err(ApiErrorResponse)?;
    
    Ok(Json(ExportXlsxResponse { success: true }))
}

async fn import_xlsx(
    Extension(api): Extension<Arc<WorkbookApi>>,
    Json(payload): Json<ImportXlsxRequest>,
) -> Result<Json<ImportXlsxResponse>, ApiErrorResponse> {
    api.import_xlsx(payload.path.into())
        .map_err(ApiErrorResponse)?;
    
    Ok(Json(ImportXlsxResponse { success: true }))
}

// Wrapper for API errors to convert them to HTTP responses
pub struct ApiErrorResponse(ApiError);

impl IntoResponse for ApiErrorResponse {
    fn into_response(self) -> Response {
        let status = StatusCode::from(self.0);
        let body = Json(error_to_json(self.0));
        (status, body).into_response()
    }
}
