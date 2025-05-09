//! examples/dcf_model.rs
//! Demonstrates a Discounted Cash Flow (DCF) model using financial functions

use ssengine_core::{Workbook, new_workbook};
use ssengine_io::write_xlsx;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating a DCF model spreadsheet...");
    
    // Create a new workbook
    let mut wb = new_workbook();
    
    // Add sheets for the DCF model
    let assumptions_id = wb.add_sheet("Assumptions".to_string())?;
    let projections_id = wb.add_sheet("Projections".to_string())?;
    let dcf_id = wb.add_sheet("DCF Valuation".to_string())?;
    
    // Set up the Assumptions sheet
    if let Some(sheet) = wb.get_sheet_mut(&assumptions_id) {
        // Headers and formatting
        sheet.set_cell(0, 0, "DCF Model Assumptions".into())?;
        
        // General assumptions
        sheet.set_cell(2, 0, "Discount Rate (WACC)".into())?;
        sheet.set_cell(2, 1, 0.10.into())?; // 10%
        
        sheet.set_cell(3, 0, "Perpetual Growth Rate".into())?;
        sheet.set_cell(3, 1, 0.02.into())?; // 2%
        
        sheet.set_cell(4, 0, "Initial Investment".into())?;
        sheet.set_cell(4, 1, (-1000000.0).into())?; // $1M initial investment
        
        sheet.set_cell(5, 0, "Projection Years".into())?;
        sheet.set_cell(5, 1, 5.0.into())?; // 5-year projection
        
        // Revenue growth assumptions
        sheet.set_cell(7, 0, "Revenue Growth Rate".into())?;
        sheet.set_cell(7, 1, 0.15.into())?; // 15% in year 1
        sheet.set_cell(7, 2, 0.12.into())?; // 12% in year 2
        sheet.set_cell(7, 3, 0.10.into())?; // 10% in year 3
        sheet.set_cell(7, 4, 0.08.into())?; // 8% in year 4
        sheet.set_cell(7, 5, 0.05.into())?; // 5% in year 5
        
        // Margin assumptions
        sheet.set_cell(8, 0, "EBITDA Margin".into())?;
        sheet.set_cell(8, 1, 0.25.into())?; // 25% EBITDA margin 
        
        sheet.set_cell(9, 0, "Tax Rate".into())?;
        sheet.set_cell(9, 1, 0.21.into())?; // 21% tax rate
        
        sheet.set_cell(10, 0, "Capex (% of Revenue)".into())?;
        sheet.set_cell(10, 1, 0.10.into())?; // 10% of revenue
        
        sheet.set_cell(11, 0, "Working Capital (% of Revenue)".into())?;
        sheet.set_cell(11, 1, 0.15.into())?; // 15% of revenue
        
        sheet.set_cell(12, 0, "Base Year Revenue".into())?;
        sheet.set_cell(12, 1, 1000000.0.into())?; // $1M starting revenue
    }
    
    // Set up the Projections sheet
    if let Some(sheet) = wb.get_sheet_mut(&projections_id) {
        // Headers
        sheet.set_cell(0, 0, "DCF Model Projections".into())?;
        
        // Year headers
        sheet.set_cell(2, 0, "Year".into())?;
        sheet.set_cell(2, 1, "0 (Base)".into())?;
        sheet.set_cell(2, 2, "1".into())?;
        sheet.set_cell(2, 3, "2".into())?;
        sheet.set_cell(2, 4, "3".into())?;
        sheet.set_cell(2, 5, "4".into())?;
        sheet.set_cell(2, 6, "5".into())?;
        
        // Initial investment
        sheet.set_cell(3, 0, "Initial Investment".into())?;
        sheet.set_cell(3, 1, "=Assumptions!B5".into())?; // Link to assumptions sheet
        
        // Revenue
        sheet.set_cell(4, 0, "Revenue".into())?;
        sheet.set_cell(4, 1, "=Assumptions!B13".into())?; // Base year revenue
        sheet.set_cell(4, 2, "=B5*(1+Assumptions!B8)".into())?; // Year 1 revenue with growth
        sheet.set_cell(4, 3, "=C5*(1+Assumptions!C8)".into())?; // Year 2 revenue with growth
        sheet.set_cell(4, 4, "=D5*(1+Assumptions!D8)".into())?; // Year 3 revenue with growth
        sheet.set_cell(4, 5, "=E5*(1+Assumptions!E8)".into())?; // Year 4 revenue with growth
        sheet.set_cell(4, 6, "=F5*(1+Assumptions!F8)".into())?; // Year 5 revenue with growth
        
        // EBITDA
        sheet.set_cell(5, 0, "EBITDA".into())?;
        sheet.set_cell(5, 1, "=B5*Assumptions!B9".into())?;
        sheet.set_cell(5, 2, "=C5*Assumptions!B9".into())?;
        sheet.set_cell(5, 3, "=D5*Assumptions!B9".into())?;
        sheet.set_cell(5, 4, "=E5*Assumptions!B9".into())?;
        sheet.set_cell(5, 5, "=F5*Assumptions!B9".into())?;
        sheet.set_cell(5, 6, "=G5*Assumptions!B9".into())?;
        
        // Taxes
        sheet.set_cell(6, 0, "Taxes".into())?;
        sheet.set_cell(6, 1, "=B6*Assumptions!B10".into())?;
        sheet.set_cell(6, 2, "=C6*Assumptions!B10".into())?;
        sheet.set_cell(6, 3, "=D6*Assumptions!B10".into())?;
        sheet.set_cell(6, 4, "=E6*Assumptions!B10".into())?;
        sheet.set_cell(6, 5, "=F6*Assumptions!B10".into())?;
        sheet.set_cell(6, 6, "=G6*Assumptions!B10".into())?;
        
        // Capital Expenditures
        sheet.set_cell(7, 0, "Capital Expenditures".into())?;
        sheet.set_cell(7, 1, "=B5*Assumptions!B11".into())?;
        sheet.set_cell(7, 2, "=C5*Assumptions!B11".into())?;
        sheet.set_cell(7, 3, "=D5*Assumptions!B11".into())?;
        sheet.set_cell(7, 4, "=E5*Assumptions!B11".into())?;
        sheet.set_cell(7, 5, "=F5*Assumptions!B11".into())?;
        sheet.set_cell(7, 6, "=G5*Assumptions!B11".into())?;
        
        // Changes in Working Capital
        sheet.set_cell(8, 0, "Changes in Working Capital".into())?;
        sheet.set_cell(8, 2, "=(C5-B5)*Assumptions!B12".into())?; // Year 1
        sheet.set_cell(8, 3, "=(D5-C5)*Assumptions!B12".into())?; // Year 2
        sheet.set_cell(8, 4, "=(E5-D5)*Assumptions!B12".into())?; // Year 3
        sheet.set_cell(8, 5, "=(F5-E5)*Assumptions!B12".into())?; // Year 4
        sheet.set_cell(8, 6, "=(G5-F5)*Assumptions!B12".into())?; // Year 5
        
        // Free Cash Flow
        sheet.set_cell(9, 0, "Free Cash Flow".into())?;
        sheet.set_cell(9, 1, "=B4".into())?; // Initial investment in base year
        sheet.set_cell(9, 2, "=C6-C7-C8-C9".into())?; // Year 1 FCF
        sheet.set_cell(9, 3, "=D6-D7-D8-D9".into())?; // Year 2 FCF
        sheet.set_cell(9, 4, "=E6-E7-E8-E9".into())?; // Year 3 FCF
        sheet.set_cell(9, 5, "=F6-F7-F8-F9".into())?; // Year 4 FCF
        sheet.set_cell(9, 6, "=G6-G7-G8-G9".into())?; // Year 5 FCF
    }
    
    // Set up the DCF Valuation sheet
    if let Some(sheet) = wb.get_sheet_mut(&dcf_id) {
        // Headers
        sheet.set_cell(0, 0, "DCF Valuation Summary".into())?;
        
        // NPV calculation
        sheet.set_cell(2, 0, "Discount Rate (WACC)".into())?;
        sheet.set_cell(2, 1, "=Assumptions!B3".into())?;
        
        sheet.set_cell(3, 0, "5-Year FCF NPV".into())?;
        sheet.set_cell(3, 1, "=NPV(B3,Projections!C10,Projections!D10,Projections!E10,Projections!F10,Projections!G10)".into())?;
        
        // Terminal value calculation
        sheet.set_cell(5, 0, "Terminal Value Calculation".into())?;
        
        sheet.set_cell(6, 0, "Year 5 FCF".into())?;
        sheet.set_cell(6, 1, "=Projections!G10".into())?;
        
        sheet.set_cell(7, 0, "Perpetual Growth Rate".into())?;
        sheet.set_cell(7, 1, "=Assumptions!B4".into())?;
        
        sheet.set_cell(8, 0, "Terminal Value".into())?;
        sheet.set_cell(8, 1, "=B7*(1+B8)/(B3-B8)".into())?; // Gordon Growth Model: TV = FCF*(1+g)/(r-g)
        
        sheet.set_cell(9, 0, "Discounted Terminal Value".into())?;
        sheet.set_cell(9, 1, "=B9/(1+B3)^5".into())?; // Discounted back 5 years
        
        // Enterprise value
        sheet.set_cell(11, 0, "Enterprise Value".into())?;
        sheet.set_cell(11, 1, "=B4+B10".into())?; // NPV of FCF + Discounted Terminal Value
        
        // Initial investment
        sheet.set_cell(12, 0, "Initial Investment".into())?;
        sheet.set_cell(12, 1, "=Projections!B4".into())?;
        
        // Net present value
        sheet.set_cell(13, 0, "Net Present Value (NPV)".into())?;
        sheet.set_cell(13, 1, "=B12+B13".into())?; // EV + Initial Investment
        
        // IRR calculation
        sheet.set_cell(15, 0, "IRR".into())?;
        sheet.set_cell(15, 1, "=IRR(Projections!B10,Projections!C10,Projections!D10,Projections!E10,Projections!F10,Projections!G10)".into())?;
    }
    
    // Export to XLSX
    let output_path = Path::new("dcf_model.xlsx");
    write_xlsx(&wb, output_path)?;
    println!("Successfully created DCF model at {}", output_path.display());
    
    Ok(())
}
