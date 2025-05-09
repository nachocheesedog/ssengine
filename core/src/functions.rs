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
        self.register("COUNT", count);
        self.register("COUNTA", counta);
        self.register("MAX", max);
        self.register("MIN", min);
        self.register("ROUND", round);
        self.register("ROUNDDOWN", rounddown);
        self.register("ROUNDUP", roundup);
        self.register("SQRT", sqrt);
        self.register("ABS", abs);
        self.register("POWER", power);
        self.register("PRODUCT", product);
        
        // Statistical functions
        self.register("STDEV", stdev);
        self.register("STDEVP", stdevp);
        self.register("VAR", var_func);
        self.register("VARP", varp);
        self.register("MEDIAN", median);
        self.register("PERCENTILE", percentile);
        
        // Logical functions
        self.register("IF", if_func);
        self.register("AND", and);
        self.register("OR", or);
        self.register("NOT", not);
        self.register("TRUE", true_func);
        self.register("FALSE", false_func);
        self.register("ISBLANK", is_blank);
        self.register("ISERROR", is_error);
        self.register("ISNUMBER", is_number);
        
        // Text functions
        self.register("CONCATENATE", concatenate);
        self.register("LEFT", left);
        self.register("RIGHT", right);
        self.register("MID", mid);
        self.register("LEN", len);
        self.register("LOWER", lower);
        self.register("UPPER", upper);
        self.register("TRIM", trim);
        self.register("SUBSTITUTE", substitute);
        self.register("FIND", find);
        self.register("TEXT", text_format);
        
        // Date functions
        self.register("TODAY", today);
        self.register("NOW", now);
        self.register("DATE", date);
        self.register("YEAR", year);
        self.register("MONTH", month);
        self.register("DAY", day);
        self.register("WEEKDAY", weekday);
        self.register("DATEDIF", datedif);
        
        // Lookup functions
        self.register("VLOOKUP", vlookup);
        self.register("HLOOKUP", hlookup);
        self.register("INDEX", index);
        self.register("MATCH", match_func);
        self.register("CHOOSE", choose);
        
        // Information functions
        self.register("ISNA", is_na);
        self.register("NA", na);
        self.register("ISERR", is_err);
        self.register("ERROR.TYPE", error_type);
        self.register("ISTEXT", is_text);
        
        // Engineering functions
        self.register("BIN2DEC", bin2dec);
        self.register("DEC2BIN", dec2bin);
        self.register("HEX2DEC", hex2dec);
        self.register("DEC2HEX", dec2hex);
        
        // Financial functions (DCF modeling)
        self.register("NPV", npv);
        self.register("IRR", irr);
        self.register("PMT", pmt);
        self.register("PV", pv);
        self.register("FV", fv);
        self.register("IPMT", ipmt);
        self.register("PPMT", ppmt);
        self.register("NPER", nper);
        self.register("RATE", rate);
        self.register("XNPV", xnpv);
        self.register("XIRR", xirr);
        self.register("DB", db);
        self.register("SLN", sln);
        self.register("SYD", syd);
        self.register("DDB", ddb);
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

// ===== MATHEMATICAL FUNCTIONS =====

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
    let mut valid_count = 0;
    let mut total = 0.0;

    for arg in args {
        match arg {
            CellValue::Number(n) => {
                total += n;
                valid_count += 1;
            },
            CellValue::Boolean(b) => {
                total += if *b { 1.0 } else { 0.0 };
                valid_count += 1;
            },
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in AVERAGE".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }

    if valid_count == 0 {
        return Err(EngineError::EvaluationError("AVERAGE requires at least one numeric value".into()));
    }

    Ok(CellValue::Number(total / valid_count as f64))
}

// COUNT function - counts number of cells with numbers
fn count(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let count = args.iter().filter(|arg| matches!(arg, CellValue::Number(_))).count();
    Ok(CellValue::Number(count as f64))
}

// COUNTA function - counts number of cells that are not empty
fn counta(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let count = args.iter().filter(|arg| !matches!(arg, CellValue::Blank)).count();
    Ok(CellValue::Number(count as f64))
}

// MAX function - returns the largest value
fn max(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut max_value = f64::NEG_INFINITY;
    let mut found_any = false;

    for arg in args {
        match arg {
            CellValue::Number(n) => {
                max_value = max_value.max(*n);
                found_any = true;
            },
            CellValue::Boolean(b) => {
                let value = if *b { 1.0 } else { 0.0 };
                max_value = max_value.max(value);
                found_any = true;
            },
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in MAX".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }

    if !found_any {
        return Err(EngineError::EvaluationError("MAX requires at least one numeric value".into()));
    }

    Ok(CellValue::Number(max_value))
}

// MIN function - returns the smallest value
fn min(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut min_value = f64::INFINITY;
    let mut found_any = false;

    for arg in args {
        match arg {
            CellValue::Number(n) => {
                min_value = min_value.min(*n);
                found_any = true;
            },
            CellValue::Boolean(b) => {
                let value = if *b { 1.0 } else { 0.0 };
                min_value = min_value.min(value);
                found_any = true;
            },
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in MIN".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }

    if !found_any {
        return Err(EngineError::EvaluationError("MIN requires at least one numeric value".into()));
    }

    Ok(CellValue::Number(min_value))
}

// ROUND function - rounds a number to a specified number of digits
fn round(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError("ROUND requires exactly 2 arguments: number and num_digits".into()));
    }
    
    // Get the number to round
    let number = extract_number(&args[0], "number")?;
    
    // Get the number of digits to round to
    let num_digits = extract_number(&args[1], "num_digits")?;
    let num_digits_int = num_digits as i32;
    
    // Calculate the rounded value
    let multiplier = 10.0_f64.powi(num_digits_int);
    let rounded = (number * multiplier).round() / multiplier;
    
    Ok(CellValue::Number(rounded))
}

// ROUNDDOWN function - rounds a number down to a specified number of digits
fn rounddown(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError("ROUNDDOWN requires exactly 2 arguments: number and num_digits".into()));
    }
    
    // Get the number to round
    let number = extract_number(&args[0], "number")?;
    
    // Get the number of digits to round to
    let num_digits = extract_number(&args[1], "num_digits")?;
    let num_digits_int = num_digits as i32;
    
    // Calculate the rounded value
    let multiplier = 10.0_f64.powi(num_digits_int);
    let sign = if number >= 0.0 { 1.0 } else { -1.0 };
    let abs_number = number.abs();
    let rounded = sign * ((abs_number * multiplier).floor() / multiplier);
    
    Ok(CellValue::Number(rounded))
}

// ROUNDUP function - rounds a number up to a specified number of digits
fn roundup(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError("ROUNDUP requires exactly 2 arguments: number and num_digits".into()));
    }
    
    // Get the number to round
    let number = extract_number(&args[0], "number")?;
    
    // Get the number of digits to round to
    let num_digits = extract_number(&args[1], "num_digits")?;
    let num_digits_int = num_digits as i32;
    
    // Calculate the rounded value
    let multiplier = 10.0_f64.powi(num_digits_int);
    let sign = if number >= 0.0 { 1.0 } else { -1.0 };
    let abs_number = number.abs();
    let rounded = sign * ((abs_number * multiplier).ceil() / multiplier);
    
    Ok(CellValue::Number(rounded))
}

// SQRT function - returns the square root of a number
fn sqrt(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("SQRT requires exactly 1 argument".into()));
    }
    
    let number = extract_number(&args[0], "number")?;
    
    if number < 0.0 {
        return Err(EngineError::EvaluationError("SQRT requires a non-negative number".into()));
    }
    
    Ok(CellValue::Number(number.sqrt()))
}

// ABS function - returns the absolute value of a number
fn abs(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ABS requires exactly 1 argument".into()));
    }
    
    let number = extract_number(&args[0], "number")?;
    Ok(CellValue::Number(number.abs()))
}

// POWER function - returns a number raised to a power
fn power(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError("POWER requires exactly 2 arguments: number and power".into()));
    }
    
    let base = extract_number(&args[0], "base")?;
    let exponent = extract_number(&args[1], "exponent")?;
    
    // Handle negative base with non-integer exponent (which would result in a complex number)
    if base < 0.0 && exponent.fract() != 0.0 {
        return Err(EngineError::EvaluationError("POWER cannot raise a negative number to a non-integer power".into()));
    }
    
    // Handle base of zero with negative exponent (division by zero)
    if base == 0.0 && exponent < 0.0 {
        return Err(EngineError::EvaluationError("POWER cannot raise zero to a negative power".into()));
    }
    
    Ok(CellValue::Number(base.powf(exponent)))
}

// PRODUCT function - multiplies all the numbers given as arguments
fn product(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut product = 1.0;
    let mut found_any = false;

    for arg in args {
        match arg {
            CellValue::Number(n) => {
                product *= n;
                found_any = true;
            },
            CellValue::Boolean(b) => {
                if *b { product *= 1.0; } else { product *= 0.0; }
                found_any = true;
            },
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in PRODUCT".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }

    if !found_any {
        return Err(EngineError::EvaluationError("PRODUCT requires at least one numeric value".into()));
    }

    Ok(CellValue::Number(product))
}

// ===== STATISTICAL FUNCTIONS =====

// STDEV function - calculates standard deviation based on a sample
fn stdev(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 {
        return Err(EngineError::EvaluationError("STDEV requires at least 2 values".into()));
    }
    
    // Extract numeric values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in STDEV".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.len() < 2 {
        return Err(EngineError::EvaluationError("STDEV requires at least 2 numeric values".into()));
    }
    
    // Calculate mean
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    
    // Calculate sum of squared differences
    let sum_sq_diff = values.iter().map(|x| (*x - mean).powi(2)).sum::<f64>();
    
    // Calculate standard deviation (sample)
    let stdev = (sum_sq_diff / (n - 1.0)).sqrt();
    
    Ok(CellValue::Number(stdev))
}

// STDEVP function - calculates standard deviation based on the entire population
fn stdevp(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 {
        return Err(EngineError::EvaluationError("STDEVP requires at least 1 value".into()));
    }
    
    // Extract numeric values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in STDEVP".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.is_empty() {
        return Err(EngineError::EvaluationError("STDEVP requires at least 1 numeric value".into()));
    }
    
    // Calculate mean
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    
    // Calculate sum of squared differences
    let sum_sq_diff = values.iter().map(|x| (*x - mean).powi(2)).sum::<f64>();
    
    // Calculate standard deviation (population)
    let stdevp = (sum_sq_diff / n).sqrt();
    
    Ok(CellValue::Number(stdevp))
}

// VAR function - calculates variance based on a sample
fn var_func(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 {
        return Err(EngineError::EvaluationError("VAR requires at least 2 values".into()));
    }
    
    // Extract numeric values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in VAR".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.len() < 2 {
        return Err(EngineError::EvaluationError("VAR requires at least 2 numeric values".into()));
    }
    
    // Calculate mean
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    
    // Calculate sum of squared differences
    let sum_sq_diff = values.iter().map(|x| (*x - mean).powi(2)).sum::<f64>();
    
    // Calculate variance (sample)
    let variance = sum_sq_diff / (n - 1.0);
    
    Ok(CellValue::Number(variance))
}

// VARP function - calculates variance based on the entire population
fn varp(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 {
        return Err(EngineError::EvaluationError("VARP requires at least 1 value".into()));
    }
    
    // Extract numeric values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in VARP".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.is_empty() {
        return Err(EngineError::EvaluationError("VARP requires at least 1 numeric value".into()));
    }
    
    // Calculate mean
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    
    // Calculate sum of squared differences
    let sum_sq_diff = values.iter().map(|x| (*x - mean).powi(2)).sum::<f64>();
    
    // Calculate variance (population)
    let variance = sum_sq_diff / n;
    
    Ok(CellValue::Number(variance))
}

// MEDIAN function - returns the median (middle value) of the given numbers
fn median(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.is_empty() {
        return Err(EngineError::EvaluationError("MEDIAN requires at least 1 value".into()));
    }
    
    // Extract numeric values
    let mut values = Vec::new();
    for arg in args {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in MEDIAN".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.is_empty() {
        return Err(EngineError::EvaluationError("MEDIAN requires at least 1 numeric value".into()));
    }
    
    // Sort the values
    values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    
    // Find the median
    let n = values.len();
    let median = if n % 2 == 0 {
        // Even number of elements, average the two middle values
        (values[n/2 - 1] + values[n/2]) / 2.0
    } else {
        // Odd number of elements, return the middle value
        values[n/2]
    };
    
    Ok(CellValue::Number(median))
}

// PERCENTILE function - returns the k-th percentile of values in a range
fn percentile(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 {
        return Err(EngineError::EvaluationError("PERCENTILE requires at least 2 arguments: array and k".into()));
    }
    
    // Extract k (percentile value between 0 and 1)
    let k = match args.last() {
        Some(CellValue::Number(n)) => *n,
        _ => return Err(EngineError::EvaluationError("PERCENTILE requires k to be a number between 0 and 1".into())),
    };
    
    if k < 0.0 || k > 1.0 {
        return Err(EngineError::EvaluationError("PERCENTILE requires k to be between 0 and 1".into()));
    }
    
    // Extract numeric values (excluding the last argument which is k)
    let mut values = Vec::new();
    for arg in args.iter().take(args.len() - 1) {
        match arg {
            CellValue::Number(n) => values.push(*n),
            CellValue::Boolean(b) => values.push(if *b { 1.0 } else { 0.0 }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot include formulas directly in PERCENTILE".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    if values.is_empty() {
        return Err(EngineError::EvaluationError("PERCENTILE requires at least 1 numeric value in the array".into()));
    }
    
    // Sort the values
    values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    
    // Calculate the position
    let n = values.len() as f64;
    let position = k * (n - 1.0);
    let position_index = position.floor() as usize;
    let position_fraction = position - position.floor();
    
    // Calculate the percentile value
    let percentile = if position_index >= values.len() - 1 {
        values[values.len() - 1]
    } else {
        values[position_index] + position_fraction * (values[position_index + 1] - values[position_index])
    };
    
    Ok(CellValue::Number(percentile))
}

// ===== LOGICAL FUNCTIONS =====

// IF function
fn if_func(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("IF requires exactly 3 arguments: condition, value_if_true, value_if_false".into()));
    }
    
    let condition = match &args[0] {
        CellValue::Boolean(b) => *b,
        CellValue::Number(n) => *n != 0.0,
        CellValue::Text(t) => !t.is_empty(),
        CellValue::Blank => false,
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula as a condition".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    if condition {
        Ok(args[1].clone())
    } else {
        Ok(args[2].clone())
    }
}

// AND function - returns TRUE if all arguments are TRUE
fn and(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.is_empty() {
        return Err(EngineError::EvaluationError("AND requires at least one argument".into()));
    }
    
    for arg in args {
        let value = match arg {
            CellValue::Boolean(b) => *b,
            CellValue::Number(n) => *n != 0.0,
            CellValue::Text(t) => !t.is_empty(),
            CellValue::Blank => false,
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in AND".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        };
        
        if !value {
            return Ok(CellValue::Boolean(false));
        }
    }
    
    Ok(CellValue::Boolean(true))
}

// OR function - returns TRUE if any argument is TRUE
fn or(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.is_empty() {
        return Err(EngineError::EvaluationError("OR requires at least one argument".into()));
    }
    
    for arg in args {
        let value = match arg {
            CellValue::Boolean(b) => *b,
            CellValue::Number(n) => *n != 0.0,
            CellValue::Text(t) => !t.is_empty(),
            CellValue::Blank => false,
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in OR".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        };
        
        if value {
            return Ok(CellValue::Boolean(true));
        }
    }
    
    Ok(CellValue::Boolean(false))
}

// NOT function - reverses the logical value of its argument
fn not(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("NOT requires exactly one argument".into()));
    }
    
    let value = match &args[0] {
        CellValue::Boolean(b) => *b,
        CellValue::Number(n) => *n != 0.0,
        CellValue::Text(t) => !t.is_empty(),
        CellValue::Blank => false,
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in NOT".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    Ok(CellValue::Boolean(!value))
}

// TRUE function - returns the logical value TRUE
fn true_func(_args: &[CellValue]) -> Result<CellValue, EngineError> {
    Ok(CellValue::Boolean(true))
}

// FALSE function - returns the logical value FALSE
fn false_func(_args: &[CellValue]) -> Result<CellValue, EngineError> {
    Ok(CellValue::Boolean(false))
}

// ISBLANK function - returns TRUE if value is blank
fn is_blank(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISBLANK requires exactly one argument".into()));
    }
    
    match &args[0] {
        CellValue::Blank => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// ISERROR function - returns TRUE if value is an error
fn is_error(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISERROR requires exactly one argument".into()));
    }
    
    match &args[0] {
        CellValue::Error(_) => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// ISNUMBER function - returns TRUE if value is a number
fn is_number(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISNUMBER requires exactly one argument".into()));
    }
    
    match &args[0] {
        CellValue::Number(_) => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// ===== TEXT FUNCTIONS =====

// CONCATENATE function - joins several text strings into one text string
fn concatenate(args: &[CellValue]) -> Result<CellValue, EngineError> {
    let mut result = String::new();
    
    for arg in args {
        match arg {
            CellValue::Text(t) => result.push_str(t),
            CellValue::Number(n) => result.push_str(&n.to_string()),
            CellValue::Boolean(b) => result.push_str(if *b { "TRUE" } else { "FALSE" }),
            CellValue::Blank => {},
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in CONCATENATE".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        }
    }
    
    Ok(CellValue::Text(result))
}

// LEFT function - returns the first character or characters in a text string
fn left(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError("LEFT requires 1 or 2 arguments: text and [num_chars]".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in LEFT".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let num_chars = if args.len() == 2 {
        match &args[1] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("LEFT's second argument must be a number".into())),
        }
    } else {
        1 // Default is 1 character
    };
    
    let chars: Vec<char> = text.chars().collect();
    let result: String = chars.iter().take(num_chars).collect();
    
    Ok(CellValue::Text(result))
}

// RIGHT function - returns the last character or characters in a text string
fn right(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError("RIGHT requires 1 or 2 arguments: text and [num_chars]".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in RIGHT".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let num_chars = if args.len() == 2 {
        match &args[1] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("RIGHT's second argument must be a number".into())),
        }
    } else {
        1 // Default is 1 character
    };
    
    let chars: Vec<char> = text.chars().collect();
    let len = chars.len();
    let start = if num_chars >= len { 0 } else { len - num_chars };
    let result: String = chars.iter().skip(start).collect();
    
    Ok(CellValue::Text(result))
}

// MID function - returns a specific number of characters from a text string
fn mid(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("MID requires exactly 3 arguments: text, start_num, and num_chars".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in MID".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let start_num = match &args[1] {
        CellValue::Number(n) => *n as usize,
        _ => return Err(EngineError::EvaluationError("MID's second argument (start_num) must be a number".into())),
    };
    
    let num_chars = match &args[2] {
        CellValue::Number(n) => *n as usize,
        _ => return Err(EngineError::EvaluationError("MID's third argument (num_chars) must be a number".into())),
    };
    
    // Excel uses 1-based indexing for the start position
    if start_num < 1 {
        return Err(EngineError::EvaluationError("MID's start_num must be at least 1".into()));
    }
    
    let chars: Vec<char> = text.chars().collect();
    let start_index = start_num - 1; // Convert to 0-based indexing
    
    let result = if start_index >= chars.len() {
        "".to_string() // If start is beyond the length, return empty string
    } else {
        chars.iter().skip(start_index).take(num_chars).collect()
    };
    
    Ok(CellValue::Text(result))
}

// LEN function - returns the number of characters in a text string
fn len(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("LEN requires exactly one argument".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in LEN".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    Ok(CellValue::Number(text.chars().count() as f64))
}

// LOWER function - converts text to lowercase
fn lower(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("LOWER requires exactly one argument".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in LOWER".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    Ok(CellValue::Text(text.to_lowercase()))
}

// UPPER function - converts text to uppercase
fn upper(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("UPPER requires exactly one argument".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in UPPER".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    Ok(CellValue::Text(text.to_uppercase()))
}
