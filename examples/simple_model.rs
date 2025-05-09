//! examples/simple_model.rs
//! Creates a simple financial model and exports it to XLSX

use ssengine_core::{Workbook, new_workbook};
use ssengine_io::write_xlsx;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a new workbook
    let mut wb = new_workbook();
    
    // Add a sheet
    let sheet_id = wb.add_sheet("Income Statement".to_string())?;
    let sheet = wb.get_sheet_mut(&sheet_id).unwrap();
    
    // Set header
    sheet.set_cell(0, 0, "Item".into())?;
    sheet.set_cell(0, 1, "2025".into())?;
    sheet.set_cell(0, 2, "2026".into())?;
    
    // Revenue
    sheet.set_cell(1, 0, "Revenue".into())?;
    sheet.set_cell(1, 1, 1000.0.into())?;
    sheet.set_cell(1, 2, 1150.0.into())?; // 15% growth
    
    // COGS (60% of revenue)
    sheet.set_cell(2, 0, "Cost of Goods Sold".into())?;
    sheet.set_cell(2, 1, "=B2*0.6".into())?; // Formula that calculates 60% of revenue
    sheet.set_cell(2, 2, "=C2*0.6".into())?;
    
    // Gross Profit
    sheet.set_cell(3, 0, "Gross Profit".into())?;
    sheet.set_cell(3, 1, "=B2-B3".into())?; // Revenue - COGS
    sheet.set_cell(3, 2, "=C2-C3".into())?;
    
    // Operating Expenses
    sheet.set_cell(4, 0, "Operating Expenses".into())?;
    sheet.set_cell(4, 1, 250.0.into())?;
    sheet.set_cell(4, 2, 275.0.into())?;
    
    // Operating Income
    sheet.set_cell(5, 0, "Operating Income".into())?;
    sheet.set_cell(5, 1, "=B4-B5".into())?; // Gross Profit - OpEx
    sheet.set_cell(5, 2, "=C4-C5".into())?;
    
    // Tax (25%)
    sheet.set_cell(6, 0, "Taxes (25%)".into())?;
    sheet.set_cell(6, 1, "=B6*0.25".into())?; // 25% of Operating Income
    sheet.set_cell(6, 2, "=C6*0.25".into())?;
    
    // Net Income
    sheet.set_cell(7, 0, "Net Income".into())?;
    sheet.set_cell(7, 1, "=B6-B7".into())?; // Operating Income - Taxes
    sheet.set_cell(7, 2, "=C6-C7".into())?;
    
    // Export the workbook to XLSX
    write_xlsx(&wb, "simple_model.xlsx")?;
    println!("Successfully created and exported simple_model.xlsx");
    
    Ok(())
}
