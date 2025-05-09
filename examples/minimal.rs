//! examples/minimal.rs
//! A minimal example to test our spreadsheet engine

use ssengine_core::{Workbook, new_workbook};
use ssengine_io::write_xlsx;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating a minimal spreadsheet...");
    
    // Create a new workbook
    let mut wb = new_workbook();
    
    // Add a sheet
    let sheet_id = wb.add_sheet("Sheet1".to_string())?;
    
    // Get a mutable reference to the sheet
    if let Some(sheet) = wb.get_sheet_mut(&sheet_id) {
        // Set some basic content
        sheet.set_cell(0, 0, "Hello".into())?;
        sheet.set_cell(0, 1, "World".into())?;
        sheet.set_cell(1, 0, 42.0.into())?;
        sheet.set_cell(1, 1, true.into())?;
        
        // A simple formula (though evaluation won't work yet)
        sheet.set_cell(2, 0, "=A2+10".into())?;
    } else {
        eprintln!("Failed to get sheet: {}", sheet_id);
        return Err("Sheet not found".into());
    }
    
    // Export to XLSX
    let output_path = Path::new("minimal.xlsx");
    write_xlsx(&wb, output_path)?;
    println!("Successfully created {}", output_path.display());
    
    Ok(())
}
