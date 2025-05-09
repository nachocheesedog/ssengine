// ssengine-core/src/functions.rs
// Registry and implementation of spreadsheet functions

use std::collections::HashMap;
use crate::model::CellValue;
use crate::error::{EngineError, CellError};

// Function signature for spreadsheet functions
pub type FunctionImpl = fn(args: &[CellValue]) -> Result<CellValue, EngineError>;

// Registry of available functions
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionImpl>,
}

impl FunctionRegistry {
    pub fn new() -> Self {
        let mut registry = FunctionRegistry {
            functions: HashMap::new(),
        };
        
        // Register built-in functions
        registry.register_defaults();
        
        registry
    }
    
    // Register a function with the registry
    pub fn register(&mut self, name: &str, implementation: FunctionImpl) {
        self.functions.insert(name.to_uppercase(), implementation);
    }
    
    // Look up a function by name
    pub fn get(&self, name: &str) -> Option<&FunctionImpl> {
        self.functions.get(&name.to_uppercase())
    }
    
    // Register all default functions
    fn register_defaults(&mut self) {
        // Math functions
        self.register("SUM", sum);
        self.register("AVERAGE", average);
        
        // Logical functions
        self.register("IF", if_func);
        
        // Financial functions (DCF modeling)
        self.register("NPV", npv);
        self.register("IRR", irr);
        self.register("PMT", pmt);
        self.register("PV", pv);
        self.register("FV", fv);
    }
}

// Function implementations

// ===== FINANCIAL FUNCTIONS FOR DCF MODELING =====

// NPV function - Net Present Value
fn npv(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 {
        return Err(EngineError::EvaluationError("NPV requires at least 2 arguments: rate and values".into()));
    }

    // Extract rate
    let rate = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("NPV first argument (rate) must be numeric".into())),
    };

    if rate <= -1.0 {
        return Err(EngineError::EvaluationError("NPV rate must be > -1".into()));
    }

    // Calculate NPV
    let mut npv = 0.0;
    for (i, arg) in args[1..].iter().enumerate() {
        match arg {
            CellValue::Number(value) => {
                // NPV formula: each cash flow is divided by (1 + rate)^period
                // Excel's NPV assumes the first cash flow is at the end of period 1
                let period = i as f64 + 1.0;
                npv += value / f64::powf(1.0 + rate, period);
            },
            CellValue::Boolean(b) => {
                let value = if *b { 1.0 } else { 0.0 };
                let period = i as f64 + 1.0;
                npv += value / f64::powf(1.0 + rate, period);
            },
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("NPV values must be numeric, not formulas".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }

    Ok(CellValue::Number(npv))
}

// IRR function - Internal Rate of Return
fn irr(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 {
        return Err(EngineError::EvaluationError("IRR requires at least 1 argument: values".into()));
    }

    // Extract guess (defaults to 0.1 or 10%)
    let guess = if args.len() >= 2 {
        match &args[1] {
            CellValue::Number(n) => *n,
            _ => 0.1, // Default guess is 10%
        }
    } else {
        0.1 // Default guess
    };

    // Extract values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => values.push(0.0),
            _ => return Err(EngineError::EvaluationError("IRR values must be numeric".into())),
        }
    }

    // Check for at least one negative and one positive cash flow
    let mut has_neg = false;
    let mut has_pos = false;
    
    for val in &values {
        if *val < 0.0 { has_neg = true; }
        if *val > 0.0 { has_pos = true; }
    }

    if !has_neg || !has_pos {
        return Err(EngineError::EvaluationError(
            "IRR requires at least one positive and one negative cash flow".into()
        ));
    }

    // Calculate IRR using Newton-Raphson method
    let max_iterations = 100;
    let precision = 1.0e-10;
    let mut rate = guess;

    // Newton-Raphson iterations
    for _ in 0..max_iterations {
        let mut f = 0.0; // NPV at current rate
        let mut df = 0.0; // Derivative of NPV

        for (i, value) in values.iter().enumerate() {
            let period = i as f64;
            let factor = (1.0 + rate).powf(period);
            f += value / factor;
            df -= period * value / ((1.0 + rate).powf(period + 1.0));
        }

        // Calculate new rate estimate
        let new_rate = rate - f / df;
        
        // Check for convergence
        if (new_rate - rate).abs() < precision {
            return Ok(CellValue::Number(new_rate));
        }
        
        rate = new_rate;
    }

    // If we get here, the algorithm didn't converge
    Err(EngineError::EvaluationError("IRR calculation failed to converge".into()))
}

// PMT function - Payment for a loan or investment
fn pmt(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "PMT requires 3-5 arguments: rate, nper, pv, [fv], [type]".into()));
    }

    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract number of periods
    let nper = extract_number(&args[1], "nper")?;
    // Extract present value
    let pv = extract_number(&args[2], "present value")?;
    // Extract future value (defaults to 0)
    let fv = if args.len() >= 4 { extract_number(&args[3], "future value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 5 { 
        let typ = extract_number(&args[4], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("PMT type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };

    // Calculate payment
    if rate == 0.0 {
        // Simple calculation when rate is zero
        return Ok(CellValue::Number(-(pv + fv) / nper));
    }

    // The Excel PMT formula
    let pmt_value = 
        (-pv * rate * (1.0 + rate).powf(nper) - fv * rate) / 
        ((1.0 + rate).powf(nper) - 1.0) / (1.0 + rate * payment_type);

    Ok(CellValue::Number(pmt_value))
}

// PV function - Present Value
fn pv(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "PV requires 3-5 arguments: rate, nper, pmt, [fv], [type]".into()));
    }

    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract number of periods
    let nper = extract_number(&args[1], "nper")?;
    // Extract payment
    let pmt = extract_number(&args[2], "payment")?;
    // Extract future value (defaults to 0)
    let fv = if args.len() >= 4 { extract_number(&args[3], "future value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 5 { 
        let typ = extract_number(&args[4], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("PV type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };

    // Calculate present value
    if rate == 0.0 {
        // Simple calculation when rate is zero
        return Ok(CellValue::Number(-pmt * nper - fv));
    }

    // The Excel PV formula
    let pv_value = 
        ((-pmt * (1.0 + rate * payment_type) * ((1.0 + rate).powf(nper) - 1.0) / rate) - fv) / 
        (1.0 + rate).powf(nper);

    Ok(CellValue::Number(pv_value))
}

// FV function - Future Value
fn fv(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "FV requires 3-5 arguments: rate, nper, pmt, [pv], [type]".into()));
    }

    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract number of periods
    let nper = extract_number(&args[1], "nper")?;
    // Extract payment
    let pmt = extract_number(&args[2], "payment")?;
    // Extract present value (defaults to 0)
    let pv = if args.len() >= 4 { extract_number(&args[3], "present value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 5 { 
        let typ = extract_number(&args[4], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("FV type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };

    // Calculate future value
    if rate == 0.0 {
        // Simple calculation when rate is zero
        let fv_value = -pv - pmt * nper;
        return Ok(CellValue::Number(fv_value));
    }

    // The Excel FV formula
    let fv_value = 
        -pv * (1.0 + rate).powf(nper) - 
        pmt * (1.0 + rate * payment_type) * ((1.0 + rate).powf(nper) - 1.0) / rate;

    Ok(CellValue::Number(fv_value))
}

// Helper function to extract a number from a CellValue
fn extract_number(value: &CellValue, name: &str) -> Result<f64, EngineError> {
    match value {
        CellValue::Number(n) => Ok(*n),
        CellValue::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        CellValue::Blank => Ok(0.0),
        CellValue::Formula(_) => Err(EngineError::EvaluationError(format!("{} must be numeric, not a formula", name))),
        CellValue::Error(e) => Err(EngineError::CellValueError(e.clone())),
        _ => Err(EngineError::EvaluationError(format!("{} must be numeric", name))),
    }
}

// ===== BASIC FUNCTIONS =====

// SUM function
fn sum(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut total = 0.0;
    
    for arg in args {
        match arg {
            CellValue::Number(n) => total += n,
            CellValue::Text(_) => return Err(EngineError::EvaluationError("Cannot sum text values".into())),
            CellValue::Boolean(b) => total += if *b { 1.0 } else { 0.0 },
            CellValue::Blank => {}, // Ignore blank cells
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    Ok(CellValue::Number(total))
}

// AVERAGE function
fn average(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut total = 0.0;
    let mut count = 0;
    
    for arg in args {
        match arg {
            CellValue::Number(n) => {
                total += n;
                count += 1;
            },
            CellValue::Boolean(b) => {
                total += if *b { 1.0 } else { 0.0 };
                count += 1;
            },
            CellValue::Text(_) => return Err(EngineError::EvaluationError("Cannot average text values".into())),
            CellValue::Blank => {}, // Ignore blank cells
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if count == 0 {
        return Err(EngineError::EvaluationError("AVERAGE requires at least one numeric value".into()));
    }
    
    Ok(CellValue::Number(total / count as f64))
}

// IF function
fn if_func(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("IF requires exactly 3 arguments".into()));
    }
    
    let condition = match &args[0] {
        CellValue::Boolean(b) => *b,
        CellValue::Number(n) => *n != 0.0,
        CellValue::Text(s) => !s.is_empty(),
        CellValue::Blank => false,
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Formulas should be evaluated before using in functions".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    if condition {
        Ok(args[1].clone())
    } else {
        Ok(args[2].clone())
    }
}
