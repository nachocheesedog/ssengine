// ssengine-core/src/model_updated.rs
// Core data structures for the spreadsheet engine

use std::collections::HashMap;
use crate::error::{EngineError, CellError};
use crate::evaluator::Evaluator;

// Basic type definitions
pub type RowId = u32;
pub type ColumnId = u32;

// Cell address (row, column)
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellAddress {
    pub row: RowId,
    pub col: ColumnId,
}

impl CellAddress {
    pub fn new(row: RowId, col: ColumnId) -> Self {
        CellAddress { row, col }
    }
    
    // Convert A1 notation to CellAddress
    pub fn from_a1(reference: &str) -> Result<Self, EngineError> {
        // Simple implementation - could be more robust
        let mut col_str = String::new();
        let mut row_str = String::new();
        
        for c in reference.chars() {
            if c.is_alphabetic() {
                col_str.push(c.to_ascii_uppercase());
            } else if c.is_numeric() {
                row_str.push(c);
            } else {
                return Err(EngineError::ParseError(format!("Invalid character in cell reference: {}", c)));
            }
        }
        
        if col_str.is_empty() || row_str.is_empty() {
            return Err(EngineError::ParseError(format!("Invalid cell reference format: {}", reference)));
        }
        
        // Convert column letters to 0-based index
        let col: ColumnId = col_str.chars().fold(0, |acc, c| {
            acc * 26 + (c as u32 - 'A' as u32 + 1)
        }) - 1;
        
        // Convert row to 0-based index
        let row: RowId = match row_str.parse::<RowId>() {
            Ok(r) => r - 1, // Excel rows are 1-based
            Err(_) => return Err(EngineError::ParseError(format!("Invalid row in cell reference: {}", row_str))),
        };
        
        Ok(CellAddress { row, col })
    }
    
    // Convert to A1 notation
    pub fn to_a1(&self) -> String {
        let mut col_str = String::new();
        let mut col_num = self.col + 1; // Convert to 1-based for conversion
        
        while col_num > 0 {
            let remainder = (col_num - 1) % 26;
            col_str.push((b'A' + remainder as u8) as char);
            col_num = (col_num - remainder) / 26;
        }
        
        format!("{}{}", col_str.chars().rev().collect::<String>(), self.row + 1)
    }
}

// Cell value enum
#[derive(Debug, Clone)]
pub enum CellValue {
    Blank,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(CellError),
    Formula(String), // The formula text
}

impl From<f64> for CellValue {
    fn from(value: f64) -> Self {
        CellValue::Number(value)
    }
}

impl From<&str> for CellValue {
    fn from(value: &str) -> Self {
        // Check if it's a formula (starts with =)
        if value.starts_with('=') {
            CellValue::Formula(value.to_string())
        } else {
            CellValue::Text(value.to_string())
        }
    }
}

impl From<String> for CellValue {
    fn from(value: String) -> Self {
        // Check if it's a formula (starts with =)
        if value.starts_with('=') {
            CellValue::Formula(value)
        } else {
            CellValue::Text(value)
        }
    }
}

impl From<bool> for CellValue {
    fn from(value: bool) -> Self {
        CellValue::Boolean(value)
    }
}

// Cell structure
#[derive(Debug, Clone)]
pub struct Cell {
    pub value: CellValue,
    pub formula: Option<String>,
    pub calculated_value: Option<CellValue>, // Result after formula evaluation
}

impl Cell {
    pub fn new(value: CellValue) -> Self {
        let formula = match &value {
            CellValue::Formula(f) => Some(f.clone()),
            _ => None,
        };
        
        Cell {
            value,
            formula,
            calculated_value: None,
        }
    }
    
    pub fn effective_value(&self) -> &CellValue {
        match &self.calculated_value {
            Some(cv) => cv,
            None => &self.value,
        }
    }
}

// A cell range (from one cell to another)
pub struct CellRange {
    pub start: CellAddress,
    pub end: CellAddress,
}

// Sheet structure
pub struct Sheet {
    name: String,
    cells: HashMap<(RowId, ColumnId), Cell>,
}

impl Sheet {
    pub fn new(name: String) -> Self {
        Sheet {
            name,
            cells: HashMap::new(),
        }
    }
    
    pub fn name(&self) -> &str {
        &self.name
    }
    
    // Get a cell at the specified coordinates
    pub fn get_cell(&self, row: RowId, col: ColumnId) -> Option<&Cell> {
        self.cells.get(&(row, col))
    }
    
    // Set a cell value at the specified coordinates
    pub fn set_cell(&mut self, row: RowId, col: ColumnId, value: CellValue) -> Result<(), EngineError> {
        let cell = Cell::new(value);
        self.cells.insert((row, col), cell);
        Ok(())
    }
    
    // Get cells in a range
    pub fn get_cell_range(&self, range: &CellRange) -> Vec<&Cell> {
        let mut result = Vec::new();
        
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                if let Some(cell) = self.get_cell(row, col) {
                    result.push(cell);
                }
            }
        }
        
        result
    }
}

// Workbook structure - the top-level container
pub struct Workbook {
    sheets: HashMap<String, Sheet>,
    active_sheet: Option<String>,
    evaluator: Evaluator,
}

impl Workbook {
    pub fn new() -> Self {
        Workbook {
            sheets: HashMap::new(),
            active_sheet: None,
            evaluator: Evaluator::new(),
        }
    }
    
    // Add a new sheet to the workbook
    pub fn add_sheet(&mut self, name: String) -> Result<String, EngineError> {
        if self.sheets.contains_key(&name) {
            return Err(EngineError::Internal(format!("Sheet '{}' already exists", name)));
        }
        
        let sheet = Sheet::new(name.clone());
        self.sheets.insert(name.clone(), sheet);
        
        // If this is the first sheet, make it active
        if self.active_sheet.is_none() {
            self.active_sheet = Some(name.clone());
        }
        
        Ok(name)
    }
    
    // Get a sheet by name
    pub fn get_sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.get(name)
    }
    
    // Get a mutable reference to a sheet
    pub fn get_sheet_mut(&mut self, name: &str) -> Option<&mut Sheet> {
        self.sheets.get_mut(name)
    }
    
    // Set the active sheet
    pub fn set_active_sheet(&mut self, name: &str) -> Result<(), EngineError> {
        if !self.sheets.contains_key(name) {
            return Err(EngineError::Internal(format!("Sheet '{}' does not exist", name)));
        }
        
        self.active_sheet = Some(name.to_string());
        Ok(())
    }
    
    // Get the active sheet name
    pub fn active_sheet_name(&self) -> Option<&String> {
        self.active_sheet.as_ref()
    }
    
    // Count sheets
    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }
    
    // Get all sheet names
    pub fn sheet_names(&self) -> Vec<&String> {
        self.sheets.keys().collect()
    }
}
