// ssengine-core/src/model.rs
// Core data structures for the spreadsheet engine

use std::collections::{HashMap, HashSet};
use std::collections::VecDeque;
use std::fmt;
use crate::error::{EngineError, CellError};
use crate::evaluator::Evaluator;
use crate::parser::Parser;

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
    
    // Get a mutable reference to a cell at the specified coordinates
    pub fn get_cell_mut(&mut self, row: RowId, col: ColumnId) -> Option<&mut Cell> {
        self.cells.get_mut(&(row, col))
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
    
    // Get cell value as a vector for a range (useful for formula evaluation)
    pub fn get_range_values(&self, range: &CellRange) -> Vec<CellValue> {
        let mut result = Vec::new();
        
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                match self.get_cell(row, col) {
                    Some(cell) => result.push(cell.effective_value().clone()),
                    None => result.push(CellValue::Blank),
                }
            }
        }
        
        result
    }
    
    // Check if the sheet contains a cell at the specified coordinates
    pub fn contains_cell(&self, row: RowId, col: ColumnId) -> bool {
        self.cells.contains_key(&(row, col))
    }
    
    // Delete a cell at the specified coordinates
    pub fn delete_cell(&mut self, row: RowId, col: ColumnId) -> bool {
        self.cells.remove(&(row, col)).is_some()
    }
    
    // Get all cells in the sheet
    pub fn get_all_cells(&self) -> impl Iterator<Item = (&(RowId, ColumnId), &Cell)> {
        self.cells.iter()
    }
    
    // Get the number of cells in the sheet
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }
    
    // Find all cells in a given row
    pub fn get_row(&self, row: RowId) -> impl Iterator<Item = (&ColumnId, &Cell)> + '_ {
        self.cells.iter()
            .filter(move |((r, _), _)| *r == row)
            .map(|((_, c), cell)| (c, cell))
    }
    
    // Find all cells in a given column
    pub fn get_column(&self, col: ColumnId) -> impl Iterator<Item = (&RowId, &Cell)> + '_ {
        self.cells.iter()
            .filter(move |((_, c), _)| *c == col)
            .map(|((r, _), cell)| (r, cell))
    }
}

// Cell reference struct for tracking dependencies
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct CellReference {
    pub sheet: Option<String>, // None means current sheet
    pub address: CellAddress,
}

impl CellReference {
    pub fn new(address: CellAddress) -> Self {
        CellReference {
            sheet: None,
            address,
        }
    }
    
    pub fn with_sheet(sheet: String, address: CellAddress) -> Self {
        CellReference {
            sheet: Some(sheet),
            address,
        }
    }
    
    pub fn from_a1(reference: &str) -> Result<Self, EngineError> {
        // Split sheet and cell reference if sheet is specified (Sheet1!A1)
        if let Some(pos) = reference.find('!') {
            let (sheet, cell) = reference.split_at(pos);
            let cell = &cell[1..]; // Remove the '!' character
            let address = CellAddress::from_a1(cell)?;
            Ok(CellReference::with_sheet(sheet.to_string(), address))
        } else {
            let address = CellAddress::from_a1(reference)?;
            Ok(CellReference::new(address))
        }
    }
}

impl fmt::Display for CellReference {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(sheet) = &self.sheet {
            write!(f, "{}!{}", sheet, self.address.to_a1())
        } else {
            write!(f, "{}", self.address.to_a1())
        }
    }
}

// Dependency graph for tracking cell references
pub struct DependencyGraph {
    // Maps cells to the cells they depend on (cell -> [dependent cells])
    dependents: HashMap<(String, CellAddress), HashSet<(String, CellAddress)>>,
    
    // Maps cells to the cells they reference (cell -> [referenced cells])
    precedents: HashMap<(String, CellAddress), HashSet<(String, CellAddress)>>,
}

impl DependencyGraph {
    pub fn new() -> Self {
        DependencyGraph {
            dependents: HashMap::new(),
            precedents: HashMap::new(),
        }
    }
    
    // Add a dependency edge: cell depends on dependency
    pub fn add_dependency(&mut self, sheet: &str, cell: &CellAddress, dep_sheet: &str, dependency: &CellAddress) {
        let cell_key = (sheet.to_string(), cell.clone());
        let dep_key = (dep_sheet.to_string(), dependency.clone());
        
        // Add to dependents map (dependency -> cell that depends on it)
        self.dependents.entry(dep_key.clone())
            .or_insert_with(HashSet::new)
            .insert(cell_key.clone());
        
        // Add to precedents map (cell -> cells it depends on)
        self.precedents.entry(cell_key)
            .or_insert_with(HashSet::new)
            .insert(dep_key);
    }
    
    // Get all cells that depend on this cell (directly or indirectly)
    pub fn get_dependents(&self, sheet: &str, cell: &CellAddress) -> HashSet<(String, CellAddress)> {
        let mut result = HashSet::new();
        let mut queue = VecDeque::new();
        
        // Start with direct dependents
        let key = (sheet.to_string(), cell.clone());
        if let Some(deps) = self.dependents.get(&key) {
            for dep in deps {
                queue.push_back(dep.clone());
                result.insert(dep.clone());
            }
        }
        
        // Process the queue to get indirect dependents
        while let Some(current) = queue.pop_front() {
            if let Some(deps) = self.dependents.get(&current) {
                for dep in deps {
                    if !result.contains(dep) {
                        queue.push_back(dep.clone());
                        result.insert(dep.clone());
                    }
                }
            }
        }
        
        result
    }
    
    // Get direct precedents (cells this cell depends on)
    pub fn get_precedents(&self, sheet: &str, cell: &CellAddress) -> Option<&HashSet<(String, CellAddress)>> {
        self.precedents.get(&(sheet.to_string(), cell.clone()))
    }
    
    // Remove all dependencies for a cell (when cell is updated/removed)
    pub fn remove_dependencies(&mut self, sheet: &str, cell: &CellAddress) {
        let cell_key = (sheet.to_string(), cell.clone());
        
        // Remove from precedents map and collect all precedents
        let precedents = self.precedents.remove(&cell_key).unwrap_or_default();
        
        // Remove this cell from the dependents list of each precedent
        for (prec_sheet, prec_cell) in precedents {
            if let Some(deps) = self.dependents.get_mut(&(prec_sheet, prec_cell)) {
                deps.remove(&cell_key);
                
                // Clean up empty sets
                if deps.is_empty() {
                    self.dependents.remove(&(prec_sheet, prec_cell));
                }
            }
        }
    }
    
    // Check for circular references starting from this cell
    pub fn check_circular_reference(&self, sheet: &str, cell: &CellAddress) -> bool {
        let mut visited = HashSet::new();
        let mut path = HashSet::new();
        
        self.dfs_check_circular(&(sheet.to_string(), cell.clone()), &mut visited, &mut path)
    }
    
    // Helper DFS function for circular reference detection
    fn dfs_check_circular(&self, 
                           cell: &(String, CellAddress), 
                           visited: &mut HashSet<(String, CellAddress)>, 
                           path: &mut HashSet<(String, CellAddress)>) -> bool {
        // If we've seen this cell in our current path, we have a cycle
        if path.contains(cell) {
            return true;
        }
        
        // If we've already visited this cell and found no cycles, skip it
        if visited.contains(cell) {
            return false;
        }
        
        // Add to current path and mark as visited
        path.insert(cell.clone());
        visited.insert(cell.clone());
        
        // Check all precedents for cycles
        if let Some(precedents) = self.precedents.get(cell) {
            for precedent in precedents {
                if self.dfs_check_circular(precedent, visited, path) {
                    return true;
                }
            }
        }
        
        // Remove from current path (backtrack)
        path.remove(cell);
        
        false
    }
}

// Workbook structure - the top-level container
pub struct Workbook {
    sheets: HashMap<String, Sheet>,
    active_sheet: Option<String>,
    evaluator: Evaluator,
    parser: Parser,
    dependency_graph: DependencyGraph,
    dirty_cells: HashSet<(String, CellAddress)>,
}

impl Workbook {
    pub fn new() -> Self {
        Workbook {
            sheets: HashMap::new(),
            active_sheet: None,
            evaluator: Evaluator::new(),
            parser: Parser::new(),
            dependency_graph: DependencyGraph::new(),
            dirty_cells: HashSet::new(),
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
    
    // Set a cell value and update dependencies
    pub fn set_cell_value(&mut self, sheet_name: &str, row: RowId, col: ColumnId, value: impl Into<CellValue>) -> Result<(), EngineError> {
        let cell_addr = CellAddress::new(row, col);
        let value = value.into();
        
        // Get sheet or return error
        let sheet = match self.sheets.get_mut(sheet_name) {
            Some(s) => s,
            None => return Err(EngineError::Internal(format!("Sheet '{}' does not exist", sheet_name))),
        };
        
        // Clear existing dependencies for this cell
        self.dependency_graph.remove_dependencies(sheet_name, &cell_addr);
        
        // If it's a formula, parse it and update dependencies
        if let CellValue::Formula(formula_text) = &value {
            // Parse formula to get AST
            let ast = self.parser.parse(formula_text)?;
            
            // Extract cell references from AST and add to dependency graph
            let references = self.parser.extract_cell_references(&ast);
            for reference in references {
                let ref_sheet = reference.sheet.as_deref().unwrap_or(sheet_name);
                self.dependency_graph.add_dependency(sheet_name, &cell_addr, ref_sheet, &reference.address);
            }
            
            // Check for circular references
            if self.dependency_graph.check_circular_reference(sheet_name, &cell_addr) {
                // Undo adding dependencies
                self.dependency_graph.remove_dependencies(sheet_name, &cell_addr);
                return Err(EngineError::CircularReference(format!("Circular reference detected at {}", cell_addr.to_a1())));
            }
        }
        
        // Set the cell value
        sheet.set_cell(row, col, value)?;
        
        // Mark this cell and its dependents as dirty
        self.mark_dirty(sheet_name, &cell_addr);
        
        // Recalculate dirty cells
        self.recalculate()?;
        
        Ok(())
    }
    
    // Mark a cell and all its dependents as dirty (needs recalculation)
    fn mark_dirty(&mut self, sheet_name: &str, cell_addr: &CellAddress) {
        self.dirty_cells.insert((sheet_name.to_string(), cell_addr.clone()));
        
        // Mark all dependent cells as dirty
        let dependents = self.dependency_graph.get_dependents(sheet_name, cell_addr);
        for (dep_sheet, dep_addr) in dependents {
            self.dirty_cells.insert((dep_sheet, dep_addr));
        }
    }
    
    // Recalculate all dirty cells
    pub fn recalculate(&mut self) -> Result<(), EngineError> {
        // Sort dirty cells in topological order 
        // (In a full implementation, we'd do a proper topological sort of the dependency graph)
        let dirty_cells = std::mem::take(&mut self.dirty_cells);
        
        for (sheet_name, cell_addr) in dirty_cells {
            // Get the sheet
            let sheet = match self.sheets.get(&sheet_name) {
                Some(s) => s,
                None => continue, // Skip if sheet doesn't exist (shouldn't happen)
            };
            
            // Get the cell
            let cell = match sheet.get_cell(cell_addr.row, cell_addr.col) {
                Some(c) => c,
                None => continue, // Skip if cell doesn't exist (shouldn't happen)
            };
            
            // Only evaluate formulas
            if let CellValue::Formula(formula) = &cell.value {
                // Clone the formula (because we need to pass ownership to evaluate)
                let formula_clone = formula.clone();
                
                // Evaluate the formula
                let result = self.evaluator.evaluate_formula(self, &sheet_name, &cell_addr, &formula_clone)?;
                
                // Update the calculated value
                let sheet = self.sheets.get_mut(&sheet_name).unwrap();
                let cell = sheet.get_cell_mut(cell_addr.row, cell_addr.col).unwrap();
                cell.calculated_value = Some(result);
            }
        }
        
        Ok(())
    }
    
    // Get a cell value (calculated value if formula, or direct value)
    pub fn get_cell_value(&self, sheet_name: &str, row: RowId, col: ColumnId) -> Result<CellValue, EngineError> {
        let sheet = match self.get_sheet(sheet_name) {
            Some(s) => s,
            None => return Err(EngineError::Internal(format!("Sheet '{}' does not exist", sheet_name))),
        };
        
        let cell = match sheet.get_cell(row, col) {
            Some(c) => c,
            None => return Ok(CellValue::Blank), // Empty cell is treated as blank
        };
        
        // Return calculated value if available, otherwise the direct value
        match &cell.calculated_value {
            Some(cv) => Ok(cv.clone()),
            None => Ok(cell.value.clone()),
        }
    }
}
