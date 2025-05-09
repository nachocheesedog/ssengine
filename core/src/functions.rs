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
        self.register("MOD", mod_func);
        self.register("CEILING", ceiling);
        self.register("FLOOR", floor);
        self.register("MROUND", mround);
        self.register("TRANSPOSE", transpose);
        self.register("LOG", log_func);
        self.register("LN", ln);
        self.register("EXP", exp);
        self.register("RAND", rand);
        self.register("RANDBETWEEN", randbetween);
        self.register("RANDARRAY", randarray);
        
        // Conditional aggregates
        self.register("SUMIF", sumif);
        self.register("SUMIFS", sumifs);
        self.register("COUNTIF", countif);
        self.register("COUNTIFS", countifs);
        self.register("AVERAGEIF", averageif);
        self.register("AVERAGEIFS", averageifs);
        self.register("SUMPRODUCT", sumproduct);
        
        // Statistical functions
        self.register("STDEV", stdev);
        self.register("STDEVP", stdevp);
        self.register("VAR", var_func);
        self.register("VARP", varp);
        self.register("MEDIAN", median);
        self.register("PERCENTILE", percentile);
        self.register("MODE.SNGL", mode_sngl);
        self.register("COVARIANCE.P", covariance_p);
        self.register("CORREL", correl);
        self.register("AGGREGATE", aggregate);
        
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
        self.register("IFERROR", iferror);
        self.register("IFNA", ifna);
        self.register("IFS", ifs);
        
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
        self.register("TEXTJOIN", textjoin);
        
        // Date functions
        self.register("TODAY", today);
        self.register("NOW", now);
        self.register("DATE", date);
        self.register("YEAR", year);
        self.register("MONTH", month);
        self.register("DAY", day);
        self.register("WEEKDAY", weekday);
        self.register("DATEDIF", datedif);
        self.register("EOMONTH", eomonth);
        self.register("EDATE", edate);
        self.register("NETWORKDAYS", networkdays);
        self.register("NETWORKDAYS.INTL", networkdays_intl);
        self.register("WORKDAY", workday);
        self.register("WORKDAY.INTL", workday_intl);
        self.register("YEARFRAC", yearfrac);
        
        // Lookup functions
        self.register("VLOOKUP", vlookup);
        self.register("HLOOKUP", hlookup);
        self.register("INDEX", index);
        self.register("MATCH", match_func);
        self.register("CHOOSE", choose);
        self.register("XLOOKUP", xlookup);
        self.register("XMATCH", xmatch);
        self.register("OFFSET", offset);
        self.register("INDIRECT", indirect);
        
        // Dynamic array functions
        self.register("FILTER", filter);
        self.register("SORT", sort);
        self.register("UNIQUE", unique);
        self.register("SEQUENCE", sequence);
        self.register("LET", let_func);
        self.register("LAMBDA", lambda);
        
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
        self.register("MIRR", mirr);
        self.register("CUMIPMT", cumipmt);
        self.register("CUMPRINC", cumprinc);
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

// TRIM function - removes spaces from text except single spaces between words
fn trim(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("TRIM requires exactly one argument".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in TRIM".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    // First trim leading and trailing spaces
    let mut result = text.trim().to_string();
    
    // Replace multiple spaces with a single space
    let mut prev_char_was_space = false;
    result = result.chars().filter(|c| {
        let is_space = *c == ' ';
        let include = !is_space || !prev_char_was_space;
        prev_char_was_space = is_space;
        include
    }).collect();
    
    Ok(CellValue::Text(result))
}

// SUBSTITUTE function - substitutes new text for old text in a text string
fn substitute(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 4 {
        return Err(EngineError::EvaluationError("SUBSTITUTE requires 3 or 4 arguments: text, old_text, new_text, [instance_num]".into()));
    }
    
    let text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in SUBSTITUTE".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let old_text = match &args[1] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in SUBSTITUTE".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let new_text = match &args[2] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in SUBSTITUTE".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    // Check if we're replacing a specific instance
    if args.len() == 4 {
        let instance_num = match &args[3] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("SUBSTITUTE's fourth argument must be a number".into())),
        };
        
        if instance_num < 1 {
            return Err(EngineError::EvaluationError("SUBSTITUTE's instance_num must be at least 1".into()));
        }
        
        // Replace specific instance
        let mut current_instance = 0;
        let mut result = text.clone();
        let mut last_end = 0;
        let mut final_result = String::new();
        
        while let Some(start) = result[last_end..].find(&old_text) {
            current_instance += 1;
            let actual_start = last_end + start;
            
            if current_instance == instance_num {
                // Add everything before this instance
                final_result.push_str(&result[..actual_start]);
                // Add the replacement
                final_result.push_str(&new_text);
                // Set last_end to after this instance
                last_end = actual_start + old_text.len();
                // Add everything after this instance
                final_result.push_str(&result[last_end..]);
                return Ok(CellValue::Text(final_result));
            } else {
                last_end = actual_start + old_text.len();
            }
        }
        
        // If we get here, we didn't find enough instances, so return original text
        return Ok(CellValue::Text(text));
    } else {
        // Replace all instances
        let result = text.replace(&old_text, &new_text);
        Ok(CellValue::Text(result))
    }
}

// FIND function - finds one text string within another text string (case-sensitive)
fn find(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError("FIND requires 2 or 3 arguments: find_text, within_text, [start_num]".into()));
    }
    
    let find_text = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in FIND".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let within_text = match &args[1] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        CellValue::Blank => "".to_string(),
        CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in FIND".into())),
        CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
    };
    
    let start_num = if args.len() == 3 {
        match &args[2] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("FIND's third argument must be a number".into())),
        }
    } else {
        1 // Default is to start at the beginning (position 1)
    };
    
    if start_num < 1 {
        return Err(EngineError::EvaluationError("FIND's start_num must be at least 1".into()));
    }
    
    let within_chars: Vec<char> = within_text.chars().collect();
    if start_num > within_chars.len() {
        return Err(EngineError::EvaluationError("Start position in FIND is beyond the length of the text".into()));
    }
    
    // Adjust for 1-based indexing
    let start_index = start_num - 1;
    
    // Get substring from start_index to end
    let substring: String = within_chars.iter().skip(start_index).collect();
    
    // Find the position of find_text in the substring
    match substring.find(&find_text) {
        Some(pos) => Ok(CellValue::Number((start_index + pos + 1) as f64)), // +1 for 1-based indexing
        None => Err(EngineError::EvaluationError(format!("Cannot find '{}' within '{}'", find_text, within_text).into())),
    }
}

// TEXT function - formats a number and converts it to text
fn text_format(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError("TEXT requires exactly 2 arguments: value and format_text".into()));
    }
    
    let value = match &args[0] {
        CellValue::Number(n) => *n,
        CellValue::Boolean(b) => if *b { 1.0 } else { 0.0 },
        _ => return Err(EngineError::EvaluationError("TEXT's first argument must be a number".into())),
    };
    
    let format_text = match &args[1] {
        CellValue::Text(t) => t.clone(),
        _ => return Err(EngineError::EvaluationError("TEXT's second argument must be a text string".into())),
    };
    
    // This is a simplified implementation of TEXT
    // In a real implementation, you would parse the format_text and apply the formatting rules
    
    // Basic implementation for common formats
    if format_text == "0" {
        return Ok(CellValue::Text(format!("{:.0}", value)));
    } else if format_text == "0.00" {
        return Ok(CellValue::Text(format!("{:.2}", value)));
    } else if format_text == "#,##0" {
        let formatted = format!("{:.0}", value)
            .chars()
            .rev()
            .collect::<Vec<_>>()
            .chunks(3)
            .map(|chunk| chunk.iter().collect::<String>())
            .collect::<Vec<_>>()
            .join(",")
            .chars()
            .rev()
            .collect::<String>();
        return Ok(CellValue::Text(formatted));
    } else if format_text == "#,##0.00" {
        let mut parts = format!("{:.2}", value).split('.').collect::<Vec<_>>();
        if parts.len() == 2 {
            let whole_part = parts[0]
                .chars()
                .rev()
                .collect::<Vec<_>>()
                .chunks(3)
                .map(|chunk| chunk.iter().collect::<String>())
                .collect::<Vec<_>>()
                .join(",")
                .chars()
                .rev()
                .collect::<String>();
            return Ok(CellValue::Text(format!("{}.{}", whole_part, parts[1])));
        }
    } else if format_text == "0%" {
        return Ok(CellValue::Text(format!("{:.0}%", value * 100.0)));
    } else if format_text == "0.00%" {
        return Ok(CellValue::Text(format!("{:.2}%", value * 100.0)));
    } else if format_text == "$#,##0.00" {
        let mut parts = format!("{:.2}", value).split('.').collect::<Vec<_>>();
        if parts.len() == 2 {
            let whole_part = parts[0]
                .chars()
                .rev()
                .collect::<Vec<_>>()
                .chunks(3)
                .map(|chunk| chunk.iter().collect::<String>())
                .collect::<Vec<_>>()
                .join(",")
                .chars()
                .rev()
                .collect::<String>();
            return Ok(CellValue::Text(format!("${}.{}", whole_part, parts[1])));
        }
    }
    
    // Default fallback for unimplemented formats
    Ok(CellValue::Text(value.to_string()))
}

// ===== DATE FUNCTIONS =====

// TODAY function - returns the current date
fn today(_args: &[CellValue]) -> Result<CellValue, EngineError> {
    // In a real implementation, this would return the current date as a serial number
    // For now, we'll return the number of days since the Excel epoch (January 1, 1900)
    // This is simplified - in a real implementation, we would use a proper date library
    
    // 44000 represents a date in 2020 (approximate value)
    Ok(CellValue::Number(44000.0))
}

// NOW function - returns the current date and time
fn now(_args: &[CellValue]) -> Result<CellValue, EngineError> {
    // Similar to TODAY but includes time as a fractional part of the day
    // 44000.5 would represent noon on that day
    Ok(CellValue::Number(44000.5))
}

// DATE function - returns the serial number of a particular date
fn date(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("DATE requires exactly 3 arguments: year, month, day".into()));
    }
    
    let year = match &args[0] {
        CellValue::Number(n) => *n as i32,
        _ => return Err(EngineError::EvaluationError("DATE's year argument must be a number".into())),
    };
    
    let month = match &args[1] {
        CellValue::Number(n) => *n as i32,
        _ => return Err(EngineError::EvaluationError("DATE's month argument must be a number".into())),
    };
    
    let day = match &args[2] {
        CellValue::Number(n) => *n as i32,
        _ => return Err(EngineError::EvaluationError("DATE's day argument must be a number".into())),
    };
    
    // This is a simplified implementation
    // In a real implementation, we would convert the year, month, and day to an Excel serial number
    // We would also handle overflow/underflow of month and day values
    
    // For now, we'll just return a placeholder value
    // This should be replaced with actual calculation logic
    Ok(CellValue::Number(44000.0))
}

// YEAR function - returns the year component of a date
fn year(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("YEAR requires exactly 1 argument: serial_number".into()));
    }
    
    let serial_number = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("YEAR's argument must be a date serial number".into())),
    };
    
    // This is a simplified implementation that doesn't do actual date calculation
    // In a real implementation, we would convert the serial number to a date and extract the year
    
    // For demonstration, we'll just return 2020 for any input
    Ok(CellValue::Number(2020.0))
}

// MONTH function - returns the month component of a date
fn month(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("MONTH requires exactly 1 argument: serial_number".into()));
    }
    
    let serial_number = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("MONTH's argument must be a date serial number".into())),
    };
    
    // This is a simplified implementation
    // For demonstration, we'll just return 1 (January) for any input
    Ok(CellValue::Number(1.0))
}

// DAY function - returns the day component of a date
fn day(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("DAY requires exactly 1 argument: serial_number".into()));
    }
    
    let serial_number = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("DAY's argument must be a date serial number".into())),
    };
    
    // This is a simplified implementation
    // For demonstration, we'll just return 1 for any input
    Ok(CellValue::Number(1.0))
}

// WEEKDAY function - returns the day of the week as a number
fn weekday(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError("WEEKDAY requires 1 or 2 arguments: serial_number, [return_type]".into()));
    }
    
    let serial_number = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("WEEKDAY's first argument must be a date serial number".into())),
    };
    
    let return_type = if args.len() == 2 {
        match &args[1] {
            CellValue::Number(n) => *n as i32,
            _ => return Err(EngineError::EvaluationError("WEEKDAY's second argument must be a number".into())),
        }
    } else {
        1 // Default return_type is 1 (1 = Sunday, 2 = Monday, ..., 7 = Saturday)
    };
    
    // This is a simplified implementation
    // For demonstration, we'll just return 1 (Sunday) for any input
    Ok(CellValue::Number(1.0))
}

// DATEDIF function - calculates the difference between two dates in various units
fn datedif(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("DATEDIF requires exactly 3 arguments: start_date, end_date, unit".into()));
    }
    
    let start_date = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("DATEDIF's first argument must be a date serial number".into())),
    };
    
    let end_date = match &args[1] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("DATEDIF's second argument must be a date serial number".into())),
    };
    
    let unit = match &args[2] {
        CellValue::Text(t) => t.clone(),
        _ => return Err(EngineError::EvaluationError("DATEDIF's third argument must be a text string".into())),
    };
    
    if start_date > end_date {
        return Err(EngineError::EvaluationError("DATEDIF's start_date must not be greater than end_date".into()));
    }
    
    // This is a simplified implementation that returns placeholder values based on the unit
    match unit.as_str() {
        "Y" => Ok(CellValue::Number(1.0)), // Years
        "M" => Ok(CellValue::Number(12.0)), // Months
        "D" => Ok(CellValue::Number(365.0)), // Days
        "YM" => Ok(CellValue::Number(0.0)), // Months excluding years
        "YD" => Ok(CellValue::Number(0.0)), // Days excluding years
        "MD" => Ok(CellValue::Number(0.0)), // Days excluding months and years
        _ => Err(EngineError::EvaluationError(format!("DATEDIF unit '{}' is not valid", unit).into())),
    }
}

// ===== LOOKUP FUNCTIONS =====

// VLOOKUP function - searches for a value in the first column of a table and returns a value in the same row
fn vlookup(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 4 {
        return Err(EngineError::EvaluationError(
            "VLOOKUP requires 3 or 4 arguments: lookup_value, table_array, col_index_num, [range_lookup]".into()
        ));
    }
    
    // In a real implementation, we would need to handle ranges and cell references
    // This is a very simplified version that just returns an error for now
    
    Err(EngineError::EvaluationError("VLOOKUP requires table_array to be a cell range reference".into()))
}

// HLOOKUP function - searches for a value in the first row of a table and returns a value in the same column
fn hlookup(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 4 {
        return Err(EngineError::EvaluationError(
            "HLOOKUP requires 3 or 4 arguments: lookup_value, table_array, row_index_num, [range_lookup]".into()
        ));
    }
    
    // In a real implementation, we would need to handle ranges and cell references
    // This is a very simplified version that just returns an error for now
    
    Err(EngineError::EvaluationError("HLOOKUP requires table_array to be a cell range reference".into()))
}

// INDEX function - returns a value from a table based on row and column numbers
fn index(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "INDEX requires 2 or 3 arguments: array, row_num, [column_num]".into()
        ));
    }
    
    // In a real implementation, we would need to handle ranges and cell references
    // This is a very simplified version that just returns an error for now
    
    Err(EngineError::EvaluationError("INDEX requires array to be a cell range reference".into()))
}

// MATCH function - searches for a value in a range and returns its relative position
fn match_func(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "MATCH requires 2 or 3 arguments: lookup_value, lookup_array, [match_type]".into()
        ));
    }
    
    // In a real implementation, we would need to handle ranges and cell references
    // This is a very simplified version that just returns an error for now
    
    Err(EngineError::EvaluationError("MATCH requires lookup_array to be a cell range reference".into()))
}

// CHOOSE function - uses an index to return a value from a list of values
fn choose(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 {
        return Err(EngineError::EvaluationError("CHOOSE requires at least 2 arguments: index_num, value1, [value2], ...".into()));
    }
    
    let index_num = match &args[0] {
        CellValue::Number(n) => *n as usize,
        _ => return Err(EngineError::EvaluationError("CHOOSE's first argument must be a number".into())),
    };
    
    if index_num < 1 || index_num >= args.len() {
        return Err(EngineError::EvaluationError(format!("CHOOSE index {} out of range", index_num).into()));
    }
    
    // Return the chosen value
    Ok(args[index_num].clone())}

// ===== INFORMATION FUNCTIONS =====

// ISNA function - checks if a value is the #N/A error
fn is_na(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISNA requires exactly 1 argument".into()));
    }
    
    // In a real implementation, we would check if the value is specifically the #N/A error
    // For now, we'll just check if it's any error and assume it's #N/A
    match &args[0] {
        CellValue::Error(_) => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// NA function - returns the #N/A error value
fn na(_args: &[CellValue]) -> Result<CellValue, EngineError> {
    Ok(CellValue::Error("#N/A".to_string()))
}

// ISERR function - checks if a value is any error except #N/A
fn is_err(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISERR requires exactly 1 argument".into()));
    }
    
    // In a real implementation, we would check if the value is any error except #N/A
    // For now, we'll just assume all errors are covered by ISERR (which isn't strictly correct)
    match &args[0] {
        CellValue::Error(_) => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// ERROR.TYPE function - returns a number corresponding to an error type
fn error_type(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ERROR.TYPE requires exactly 1 argument".into()));
    }
    
    match &args[0] {
        CellValue::Error(e) => {
            // In a real implementation, we would return different numbers for different error types
            // 1 = #NULL!, 2 = #DIV/0!, 3 = #VALUE!, 4 = #REF!, 5 = #NAME?, 6 = #NUM!, 7 = #N/A, etc.
            match e.as_str() {
                "#NULL!" => Ok(CellValue::Number(1.0)),
                "#DIV/0!" => Ok(CellValue::Number(2.0)),
                "#VALUE!" => Ok(CellValue::Number(3.0)),
                "#REF!" => Ok(CellValue::Number(4.0)),
                "#NAME?" => Ok(CellValue::Number(5.0)),
                "#NUM!" => Ok(CellValue::Number(6.0)),
                "#N/A" => Ok(CellValue::Number(7.0)),
                _ => Ok(CellValue::Number(0.0)), // Unknown error type
            }
        },
        _ => Err(EngineError::EvaluationError("ERROR.TYPE requires an error value as its argument".into())),
    }
}

// ISTEXT function - checks if a value is text
fn is_text(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("ISTEXT requires exactly 1 argument".into()));
    }
    
    match &args[0] {
        CellValue::Text(_) => Ok(CellValue::Boolean(true)),
        _ => Ok(CellValue::Boolean(false)),
    }
}

// ===== ENGINEERING FUNCTIONS =====

// BIN2DEC function - converts a binary number to decimal
fn bin2dec(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("BIN2DEC requires exactly 1 argument".into()));
    }
    
    let binary = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        _ => return Err(EngineError::EvaluationError("BIN2DEC argument must be a binary number represented as text or number".into())),
    };
    
    // Validate binary string (only 0s and 1s)
    if !binary.chars().all(|c| c == '0' || c == '1') {
        return Err(EngineError::EvaluationError("BIN2DEC argument must contain only 0s and 1s".into()));
    }
    
    // Limit to 10 bits (Excel limit for BIN2DEC)
    if binary.len() > 10 {
        return Err(EngineError::EvaluationError("BIN2DEC argument cannot exceed 10 bits".into()));
    }
    
    // Convert binary to decimal
    let mut decimal = 0i32;
    for c in binary.chars() {
        decimal = decimal * 2 + if c == '1' { 1 } else { 0 };
    }
    
    // Handle two's complement for 10-bit binary numbers starting with 1
    if binary.len() == 10 && binary.starts_with('1') {
        // Subtract 2^10 to get the negative value (two's complement)
        decimal -= 1024;
    }
    
    Ok(CellValue::Number(decimal as f64))
}

// DEC2BIN function - converts a decimal number to binary
fn dec2bin(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError("DEC2BIN requires 1 or 2 arguments: number, [places]".into()));
    }
    
    let number = match &args[0] {
        CellValue::Number(n) => *n as i32,
        _ => return Err(EngineError::EvaluationError("DEC2BIN's first argument must be a number".into())),
    };
    
    // Check if the number is within the valid range for a 10-bit two's complement binary number
    if number < -512 || number > 511 {
        return Err(EngineError::EvaluationError("DEC2BIN can only convert numbers between -512 and 511".into()));
    }
    
    let places = if args.len() == 2 {
        match &args[1] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("DEC2BIN's second argument must be a number".into())),
        }
    } else {
        0 // Default: use the minimum number of characters necessary
    };
    
    // Convert to binary
    let mut binary = String::new();
    let mut value = if number < 0 {
        // For negative numbers, use two's complement representation (10 bits)
        (number + 1024) as u32
    } else {
        number as u32
    };
    
    // Build the binary representation from right to left
    while value > 0 || binary.is_empty() {
        binary.insert(0, if value % 2 == 1 { '1' } else { '0' });
        value /= 2;
    }
    
    // Pad with leading zeros if needed
    if places > 0 && binary.len() < places {
        binary = format!("{:0>width$}", binary, width = places);
    }
    
    Ok(CellValue::Text(binary))
}

// HEX2DEC function - converts a hexadecimal number to decimal
fn hex2dec(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 1 {
        return Err(EngineError::EvaluationError("HEX2DEC requires exactly 1 argument".into()));
    }
    
    let hex = match &args[0] {
        CellValue::Text(t) => t.clone(),
        CellValue::Number(n) => n.to_string(),
        _ => return Err(EngineError::EvaluationError("HEX2DEC argument must be a hexadecimal number represented as text or number".into())),
    };
    
    // Validate hex string (only 0-9, A-F)
    if !hex.chars().all(|c| c.is_digit(16)) {
        return Err(EngineError::EvaluationError("HEX2DEC argument must contain only valid hexadecimal digits (0-9, A-F)".into()));
    }
    
    // Limit to 10 hex digits (Excel limit for HEX2DEC)
    if hex.len() > 10 {
        return Err(EngineError::EvaluationError("HEX2DEC argument cannot exceed 10 hexadecimal digits".into()));
    }
    
    // Convert hex to decimal
    match i32::from_str_radix(&hex, 16) {
        Ok(decimal) => Ok(CellValue::Number(decimal as f64)),
        Err(_) => Err(EngineError::EvaluationError("Invalid hexadecimal number".into())),
    }
}

// DEC2HEX function - converts a decimal number to hexadecimal
fn dec2hex(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError("DEC2HEX requires 1 or 2 arguments: number, [places]".into()));
    }
    
    let number = match &args[0] {
        CellValue::Number(n) => *n as i32,
        _ => return Err(EngineError::EvaluationError("DEC2HEX's first argument must be a number".into())),
    };
    
    // Check if the number is within the valid range
    if number < -549_755_813_888 || number > 549_755_813_887 {
        return Err(EngineError::EvaluationError("DEC2HEX can only convert numbers between -549,755,813,888 and 549,755,813,887".into()));
    }
    
    let places = if args.len() == 2 {
        match &args[1] {
            CellValue::Number(n) => *n as usize,
            _ => return Err(EngineError::EvaluationError("DEC2HEX's second argument must be a number".into())),
        }
    } else {
        0 // Default: use the minimum number of characters necessary
    };
    
    // Convert to hexadecimal
    let hex = if number < 0 {
        // For negative numbers, use two's complement representation
        format!("{:X}", (number + (1 << 32)) as u32)
    } else {
        format!("{:X}", number as u32)
    };
    
    // Pad with leading zeros if needed
    let result = if places > 0 && hex.len() < places {
        format!("{:0>width$}", hex, width = places)
    } else {
        hex
    };
    
    Ok(CellValue::Text(result))
}

// ===== ADDITIONAL FINANCIAL FUNCTIONS =====

// IPMT function - returns the interest payment for an investment for a given period
fn ipmt(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 4 || args.len() > 6 {
        return Err(EngineError::EvaluationError(
            "IPMT requires 4-6 arguments: rate, per, nper, pv, [fv], [type]".into()));
    }
    
    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract period
    let per = extract_number(&args[1], "per")?;
    // Extract number of periods
    let nper = extract_number(&args[2], "nper")?;
    // Extract present value
    let pv = extract_number(&args[3], "present value")?;
    // Extract future value (defaults to 0)
    let fv = if args.len() >= 5 { extract_number(&args[4], "future value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 6 { 
        let typ = extract_number(&args[5], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("IPMT type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };
    
    if per < 1.0 || per > nper {
        return Err(EngineError::EvaluationError("IPMT period must be between 1 and nper".into()));
    }
    
    // Calculate payment using PMT formula
    let pmt_value = if rate == 0.0 {
        // Simple calculation when rate is zero
        -(pv + fv) / nper
    } else {
        // The Excel PMT formula
        (-pv * rate * (1.0 + rate).powf(nper) - fv * rate) / 
        ((1.0 + rate).powf(nper) - 1.0) / (1.0 + rate * payment_type)
    };
    
    // Calculate remaining balance at the beginning of period 'per'
    let remaining_balance = if rate == 0.0 {
        pv - pmt_value * (per - 1.0)
    } else {
        pv * (1.0 + rate).powf(per - 1.0) - 
        pmt_value * ((1.0 + rate).powf(per - 1.0) - 1.0) / rate * (1.0 + rate * payment_type)
    };
    
    // Calculate interest payment for period 'per'
    let interest_payment = if payment_type == 1.0 {
        // Interest for beginning-of-period payment
        if per == 1.0 {
            0.0 // No interest payment for the first period if payment is at the beginning
        } else {
            remaining_balance * rate
        }
    } else {
        // Interest for end-of-period payment
        remaining_balance * rate
    };
    
    Ok(CellValue::Number(interest_payment))
}

// PPMT function - returns the principal payment for an investment for a given period
fn ppmt(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 4 || args.len() > 6 {
        return Err(EngineError::EvaluationError(
            "PPMT requires 4-6 arguments: rate, per, nper, pv, [fv], [type]".into()));
    }
    
    // Calculate regular payment amount (PMT)
    let pmt_args = if args.len() == 4 {
        vec![args[0].clone(), args[2].clone(), args[3].clone()]
    } else if args.len() == 5 {
        vec![args[0].clone(), args[2].clone(), args[3].clone(), args[4].clone()]
    } else {
        vec![args[0].clone(), args[2].clone(), args[3].clone(), args[4].clone(), args[5].clone()]
    };
    
    let pmt_value = match pmt(&pmt_args) {
        Ok(CellValue::Number(n)) => n,
        _ => return Err(EngineError::EvaluationError("Failed to calculate PMT value for PPMT".into())),
    };
    
    // Calculate interest payment (IPMT)
    let ipmt_value = match ipmt(args) {
        Ok(CellValue::Number(n)) => n,
        _ => return Err(EngineError::EvaluationError("Failed to calculate IPMT value for PPMT".into())),
    };
    
    // Principal payment = Regular payment - Interest payment
    Ok(CellValue::Number(pmt_value - ipmt_value))
}

// NPER function - returns the number of periods for an investment
fn nper(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "NPER requires 3-5 arguments: rate, pmt, pv, [fv], [type]".into()));
    }
    
    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    
    // If rate is zero, use simplified formula
    if rate == 0.0 {
        // Extract payment
        let pmt = extract_number(&args[1], "payment")?;
        // Extract present value
        let pv = extract_number(&args[2], "present value")?;
        // Extract future value (defaults to 0)
        let fv = if args.len() >= 4 { extract_number(&args[3], "future value")? } else { 0.0 };
        
        if pmt == 0.0 {
            return Err(EngineError::EvaluationError("NPER cannot be calculated when rate=0 and pmt=0".into()));
        }
        
        // Simple formula when rate is zero: nper = -(pv + fv) / pmt
        return Ok(CellValue::Number(-(pv + fv) / pmt));
    }
    
    // Extract payment
    let pmt = extract_number(&args[1], "payment")?;
    // Extract present value
    let pv = extract_number(&args[2], "present value")?;
    // Extract future value (defaults to 0)
    let fv = if args.len() >= 4 { extract_number(&args[3], "future value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 5 { 
        let typ = extract_number(&args[4], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("NPER type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };
    
    // Adjust payment for payment timing
    let pmt_adjusted = pmt * (1.0 + rate * payment_type);
    
    // Calculate using logarithms
    if pmt_adjusted == 0.0 {
        return Err(EngineError::EvaluationError("NPER cannot be calculated with these values".into()));
    }
    
    let nper_value = (((fv * rate + pmt_adjusted) / (pv * rate + pmt_adjusted)).ln()) / rate.ln();
    
    Ok(CellValue::Number(nper_value))
}

// RATE function - returns the interest rate per period of an annuity
fn rate(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 6 {
        return Err(EngineError::EvaluationError(
            "RATE requires 3-6 arguments: nper, pmt, pv, [fv], [type], [guess]".into()));
    }
    
    // Extract number of periods
    let nper = extract_number(&args[0], "nper")?;
    // Extract payment
    let pmt = extract_number(&args[1], "payment")?;
    // Extract present value
    let pv = extract_number(&args[2], "present value")?;
    // Extract future value (defaults to 0)
    let fv = if args.len() >= 4 { extract_number(&args[3], "future value")? } else { 0.0 };
    // Extract payment type (0 = end of period, 1 = beginning of period, defaults to 0)
    let payment_type = if args.len() >= 5 { 
        let typ = extract_number(&args[4], "type")?;
        if typ != 0.0 && typ != 1.0 {
            return Err(EngineError::EvaluationError("RATE type must be 0 or 1".into()));
        }
        typ 
    } else { 
        0.0 
    };
    // Extract initial guess (defaults to 0.1)
    let guess = if args.len() >= 6 { extract_number(&args[5], "guess")? } else { 0.1 };
    
    // Newton-Raphson method to find the rate
    let max_iterations = 100;
    let precision = 1.0e-10;
    let mut rate = guess;
    
    for _ in 0..max_iterations {
        // Calculate f(rate) - the function value at current rate
        let y = if payment_type == 1.0 {
            pv * (1.0 + rate).powf(nper) + 
            pmt * (1.0 + rate) * ((1.0 + rate).powf(nper) - 1.0) / rate + fv
        } else {
            pv * (1.0 + rate).powf(nper) + 
            pmt * ((1.0 + rate).powf(nper) - 1.0) / rate + fv
        };
        
        // If we're close enough to zero, we've found our answer
        if y.abs() < precision {
            return Ok(CellValue::Number(rate));
        }
        
        // Calculate the derivative f'(rate)
        let y1 = if payment_type == 1.0 {
            nper * pv * (1.0 + rate).powf(nper - 1.0) + 
            pmt * (1.0 + rate) * nper * (1.0 + rate).powf(nper - 1.0) / rate - 
            pmt * (1.0 + rate) * ((1.0 + rate).powf(nper) - 1.0) / (rate * rate) + 
            pmt * ((1.0 + rate).powf(nper) - 1.0) / rate
        } else {
            nper * pv * (1.0 + rate).powf(nper - 1.0) + 
            pmt * nper * (1.0 + rate).powf(nper - 1.0) / rate - 
            pmt * ((1.0 + rate).powf(nper) - 1.0) / (rate * rate)
        };
        
        // Newton-Raphson formula: new_rate = rate - f(rate) / f'(rate)
        let new_rate = rate - y / y1;
        
        // Check if we've converged
        if (new_rate - rate).abs() < precision {
            return Ok(CellValue::Number(new_rate));
        }
        
        rate = new_rate;
    }
    
    // If we didn't converge after max_iterations, return an error
    Err(EngineError::EvaluationError("RATE failed to converge".into()))
}

// XNPV function - returns the net present value for a schedule of cash flows that is not necessarily periodic
fn xnpv(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("XNPV requires exactly 3 arguments: rate, values, dates".into()));
    }
    
    // Extract rate
    let rate = match &args[0] {
        CellValue::Number(n) => *n,
        _ => return Err(EngineError::EvaluationError("XNPV rate must be numeric".into())),
    };
    
    if rate <= -1.0 {
        return Err(EngineError::EvaluationError("XNPV rate must be > -1".into()));
    }
    
    // This is a simplified implementation that doesn't handle actual ranges
    // In a real implementation, we would need to handle cell ranges for values and dates
    
    // For demonstration, we'll just return a placeholder value
    // This should be replaced with actual XNPV calculation logic
    Ok(CellValue::Number(1000.0))
}

// XIRR function - returns the internal rate of return for a schedule of cash flows that is not necessarily periodic
fn xirr(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError("XIRR requires 2 or 3 arguments: values, dates, [guess]".into()));
    }
    
    // This is a simplified implementation that doesn't handle actual ranges
    // In a real implementation, we would need to handle cell ranges for values and dates
    
    // For demonstration, we'll just return a placeholder value
    // This should be replaced with actual XIRR calculation logic
    Ok(CellValue::Number(0.1)) // 10% IRR as placeholder
}

// DB function - returns the depreciation of an asset for a specified period using the fixed-declining balance method
fn db(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 4 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "DB requires 4 or 5 arguments: cost, salvage, life, period, [month]".into()));
    }
    
    // Extract cost
    let cost = extract_number(&args[0], "cost")?;
    // Extract salvage value
    let salvage = extract_number(&args[1], "salvage")?;
    // Extract life (in years)
    let life = extract_number(&args[2], "life")?;
    // Extract period
    let period = extract_number(&args[3], "period")?;
    // Extract month (defaults to 12)
    let month = if args.len() == 5 { extract_number(&args[4], "month")? } else { 12.0 };
    
    if cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life + 1.0 {
        return Err(EngineError::EvaluationError("DB arguments out of valid range".into()));
    }
    
    if month < 1.0 || month > 12.0 {
        return Err(EngineError::EvaluationError("DB month must be between 1 and 12".into()));
    }
    
    // Calculate depreciation rate
    let rate = 1.0 - (salvage / cost).powf(1.0 / life);
    
    // Calculate depreciation
    let mut book_value = cost;
    let mut depreciation = 0.0;
    
    if period <= 1.0 {
        // First year depreciation is adjusted for the starting month
        depreciation = cost * rate * month / 12.0;
    } else if period > life {
        // Last period (period = life + 1) handles remaining depreciation
        let total_prior_depreciation = cost * (1.0 - (1.0 - rate).powf(life));
        depreciation = cost - salvage - total_prior_depreciation;
        if depreciation < 0.0 {
            depreciation = 0.0;
        }
    } else {
        // Calculate depreciation for period > 1 and <= life
        // First reduce book value by first year's depreciation
        book_value -= cost * rate * month / 12.0;
        
        // Depreciate for full years until the current period
        for _ in 1..(period as usize) {
            depreciation = book_value * rate;
            book_value -= depreciation;
        }
    }
    
    Ok(CellValue::Number(depreciation))
}

// SLN function - returns the straight-line depreciation of an asset for one period
fn sln(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError("SLN requires exactly 3 arguments: cost, salvage, life".into()));
    }
    
    // Extract cost
    let cost = extract_number(&args[0], "cost")?;
    // Extract salvage value
    let salvage = extract_number(&args[1], "salvage")?;
    // Extract life (in years)
    let life = extract_number(&args[2], "life")?;
    
    if life == 0.0 {
        return Err(EngineError::EvaluationError("SLN life cannot be zero".into()));
    }
    
    // Calculate straight-line depreciation
    let depreciation = (cost - salvage) / life;
    
    Ok(CellValue::Number(depreciation))
}

// SYD function - returns the sum-of-years' digits depreciation of an asset for a specified period
fn syd(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 4 {
        return Err(EngineError::EvaluationError("SYD requires exactly 4 arguments: cost, salvage, life, period".into()));
    }
    
    // Extract cost
    let cost = extract_number(&args[0], "cost")?;
    // Extract salvage value
    let salvage = extract_number(&args[1], "salvage")?;
    // Extract life (in years)
    let life = extract_number(&args[2], "life")?;
    // Extract period
    let period = extract_number(&args[3], "period")?;
    
    if life <= 0.0 || period <= 0.0 || period > life {
        return Err(EngineError::EvaluationError("SYD arguments out of valid range".into()));
    }
    
    // Calculate sum of years' digits
    let sum_of_digits = life * (life + 1.0) / 2.0;
    
    // Calculate depreciation
    let depreciation = (cost - salvage) * (life - period + 1.0) / sum_of_digits;
    
    Ok(CellValue::Number(depreciation))
}

// DDB function - returns the depreciation of an asset for a specified period using the double-declining balance method
fn ddb(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 4 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "DDB requires 4 or 5 arguments: cost, salvage, life, period, [factor]".into()));
    }
    
    // Extract cost
    let cost = extract_number(&args[0], "cost")?;
    // Extract salvage value
    let salvage = extract_number(&args[1], "salvage")?;
    // Extract life (in years)
    let life = extract_number(&args[2], "life")?;
    // Extract period
    let period = extract_number(&args[3], "period")?;
    // Extract factor (defaults to 2)
    let factor = if args.len() == 5 { extract_number(&args[4], "factor")? } else { 2.0 };
    
    if cost <= 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life || factor <= 0.0 {
        return Err(EngineError::EvaluationError("DDB arguments out of valid range".into()));
    }
    
    // Calculate depreciation rate
    let rate = factor / life;
    
    // Calculate book value at start of period
    let mut book_value = cost;
    let mut total_depreciation = 0.0;
    
    for i in 1..(period as usize) {
        let depreciation = book_value * rate;
        if book_value - depreciation < salvage {
            total_depreciation += book_value - salvage;
            book_value = salvage;
            break;
        } else {
            total_depreciation += depreciation;
            book_value -= depreciation;
        }
    }
    
    // Calculate depreciation for the current period
    let current_depreciation = book_value * rate;
    
    // Ensure we don't depreciate below salvage value
    if book_value - current_depreciation < salvage {
        return Ok(CellValue::Number(book_value - salvage));
    }
    
    Ok(CellValue::Number(current_depreciation))
}

// MIRR function - returns the modified internal rate of return for a series of cash flows
fn mirr(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 3 {
        return Err(EngineError::EvaluationError(
            "MIRR requires exactly 3 arguments: values, finance_rate, reinvest_rate".into()));
    }
    
    // In a real implementation, we would need to handle arrays for values
    // This is a simplified version that just returns a placeholder value
    // This should be replaced with actual MIRR calculation logic
    
    // Extract finance rate
    let finance_rate = extract_number(&args[1], "finance_rate")?;
    // Extract reinvest rate
    let reinvest_rate = extract_number(&args[2], "reinvest_rate")?;
    
    if finance_rate <= -1.0 || reinvest_rate <= -1.0 {
        return Err(EngineError::EvaluationError("MIRR rates must be > -1".into()));
    }
    
    // Placeholder for MIRR calculation
    Ok(CellValue::Number(0.12)) // 12% as placeholder
}

// CUMIPMT function - returns the cumulative interest paid between start_period and end_period
fn cumipmt(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 6 {
        return Err(EngineError::EvaluationError(
            "CUMIPMT requires exactly 6 arguments: rate, nper, pv, start_period, end_period, type".into()));
    }
    
    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract number of periods
    let nper = extract_number(&args[1], "nper")?;
    // Extract present value
    let pv = extract_number(&args[2], "present value")?;
    // Extract start period
    let start_period = extract_number(&args[3], "start_period")?;
    // Extract end period
    let end_period = extract_number(&args[4], "end_period")?;
    // Extract payment type (0 = end of period, 1 = beginning of period)
    let payment_type = extract_number(&args[5], "type")?;
    
    if rate <= 0.0 || nper <= 0.0 || start_period < 1.0 || end_period < start_period || end_period > nper || (payment_type != 0.0 && payment_type != 1.0) {
        return Err(EngineError::EvaluationError("CUMIPMT arguments out of valid range".into()));
    }
    
    // Calculate regular payment (PMT)
    let pmt_args = vec![
        CellValue::Number(rate),
        CellValue::Number(nper),
        CellValue::Number(pv),
    ];
    
    let pmt_value = match pmt(&pmt_args) {
        Ok(CellValue::Number(n)) => n,
        _ => return Err(EngineError::EvaluationError("Failed to calculate PMT value for CUMIPMT".into())),
    };
    
    // Calculate cumulative interest by summing individual interest payments
    let mut cumulative_interest = 0.0;
    
    for period in (start_period as usize)..=(end_period as usize) {
        let ipmt_args = vec![
            CellValue::Number(rate),
            CellValue::Number(period as f64),
            CellValue::Number(nper),
            CellValue::Number(pv),
            CellValue::Number(0.0), // Future value = 0 
            CellValue::Number(payment_type),
        ];
        
        let interest = match ipmt(&ipmt_args) {
            Ok(CellValue::Number(n)) => n,
            _ => return Err(EngineError::EvaluationError("Failed to calculate IPMT for period ".into())),
        };
        
        cumulative_interest += interest;
    }
    
    Ok(CellValue::Number(cumulative_interest))
}

// CUMPRINC function - returns the cumulative principal paid between start_period and end_period
fn cumprinc(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 6 {
        return Err(EngineError::EvaluationError(
            "CUMPRINC requires exactly 6 arguments: rate, nper, pv, start_period, end_period, type".into()));
    }
    
    // Extract rate
    let rate = extract_number(&args[0], "rate")?;
    // Extract number of periods
    let nper = extract_number(&args[1], "nper")?;
    // Extract present value
    let pv = extract_number(&args[2], "present value")?;
    // Extract start period
    let start_period = extract_number(&args[3], "start_period")?;
    // Extract end period
    let end_period = extract_number(&args[4], "end_period")?;
    // Extract payment type (0 = end of period, 1 = beginning of period)
    let payment_type = extract_number(&args[5], "type")?;
    
    if rate <= 0.0 || nper <= 0.0 || start_period < 1.0 || end_period < start_period || end_period > nper || (payment_type != 0.0 && payment_type != 1.0) {
        return Err(EngineError::EvaluationError("CUMPRINC arguments out of valid range".into()));
    }
    
    // Calculate regular payment (PMT)
    let pmt_args = vec![
        CellValue::Number(rate),
        CellValue::Number(nper),
        CellValue::Number(pv),
    ];
    
    let pmt_value = match pmt(&pmt_args) {
        Ok(CellValue::Number(n)) => n,
        _ => return Err(EngineError::EvaluationError("Failed to calculate PMT value for CUMPRINC".into())),
    };
    
    // Calculate cumulative principal by summing individual principal payments
    let mut cumulative_principal = 0.0;
    
    for period in (start_period as usize)..=(end_period as usize) {
        let ppmt_args = vec![
            CellValue::Number(rate),
            CellValue::Number(period as f64),
            CellValue::Number(nper),
            CellValue::Number(pv),
            CellValue::Number(0.0), // Future value = 0 
            CellValue::Number(payment_type),
        ];
        
        let principal = match ppmt(&ppmt_args) {
            Ok(CellValue::Number(n)) => n,
            _ => return Err(EngineError::EvaluationError("Failed to calculate PPMT for period ".into())),
        };
        
        cumulative_principal += principal;
    }
    
    Ok(CellValue::Number(cumulative_principal))
}

// ===== CONDITIONAL AGGREGATE FUNCTIONS =====

// SUMIF function - sums cells that meet criteria
fn sumif(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "SUMIF requires 2 or 3 arguments: range, criteria, [sum_range]".into()));
    }
    
    // In a real implementation, we would need to handle ranges
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would check criteria against range
    // and sum corresponding values in sum_range
    Ok(CellValue::Number(100.0)) // Placeholder
}

// SUMIFS function - sums cells that meet multiple criteria
fn sumifs(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() % 2 == 0 {
        return Err(EngineError::EvaluationError(
            "SUMIFS requires at least 3 arguments: sum_range, criteria_range1, criteria1, ...".into()));
    }
    
    // In a real implementation, we would need to handle multiple ranges and criteria
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would check all criteria
    // against their respective ranges and sum values in sum_range where all criteria match
    Ok(CellValue::Number(50.0)) // Placeholder
}

// COUNTIF function - counts cells that meet criteria
fn countif(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError(
            "COUNTIF requires exactly 2 arguments: range, criteria".into()));
    }
    
    // In a real implementation, we would need to handle ranges
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would count cells in range that match criteria
    Ok(CellValue::Number(5.0)) // Placeholder
}

// COUNTIFS function - counts cells that meet multiple criteria
fn countifs(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() % 2 != 0 {
        return Err(EngineError::EvaluationError(
            "COUNTIFS requires at least 2 arguments and must have an even number: criteria_range1, criteria1, ...".into()));
    }
    
    // In a real implementation, we would need to handle multiple ranges and criteria
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would count cells where all criteria match
    Ok(CellValue::Number(3.0)) // Placeholder
}

// AVERAGEIF function - averages cells that meet criteria
fn averageif(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "AVERAGEIF requires 2 or 3 arguments: range, criteria, [average_range]".into()));
    }
    
    // In a real implementation, we would need to handle ranges
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would average values in average_range
    // where corresponding cells in range meet the criteria
    Ok(CellValue::Number(20.0)) // Placeholder
}

// AVERAGEIFS function - averages cells that meet multiple criteria
fn averageifs(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() % 2 == 0 {
        return Err(EngineError::EvaluationError(
            "AVERAGEIFS requires at least 3 arguments: average_range, criteria_range1, criteria1, ...".into()));
    }
    
    // In a real implementation, we would need to handle multiple ranges and criteria
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would average values in average_range
    // where all criteria match
    Ok(CellValue::Number(15.0)) // Placeholder
}

// SUMPRODUCT function - multiplies corresponding components in arrays, then sums
fn sumproduct(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.is_empty() {
        return Err(EngineError::EvaluationError(
            "SUMPRODUCT requires at least one array".into()));
    }
    
    // In a real implementation, we would need to handle arrays
    // This is a simplified version that only handles the current arguments
    
    // Mock implementation - in a real scenario we would multiply corresponding elements
    // of arrays together and sum the products
    Ok(CellValue::Number(150.0)) // Placeholder
}

// ===== ERROR HANDLING FUNCTIONS =====

// IFERROR function - returns a value if expression is error, otherwise returns expression
fn iferror(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError(
            "IFERROR requires exactly 2 arguments: value, value_if_error".into()));
    }
    
    // If first argument is an error, return second argument
    match &args[0] {
        CellValue::Error(_) => Ok(args[1].clone()),
        _ => Ok(args[0].clone()),
    }
}

// IFNA function - returns a value if expression is #N/A, otherwise returns expression
fn ifna(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError(
            "IFNA requires exactly 2 arguments: value, value_if_na".into()));
    }
    
    // If first argument is #N/A error, return second argument
    match &args[0] {
        CellValue::Error(e) if e == "#N/A" => Ok(args[1].clone()),
        _ => Ok(args[0].clone()),
    }
}

// IFS function - checks multiple conditions and returns first matching value
fn ifs(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() % 2 != 0 {
        return Err(EngineError::EvaluationError(
            "IFS requires at least 2 arguments and must have an even number: condition1, value1, condition2, value2, ...".into()));
    }
    
    // Check each condition in sequence
    for i in (0..args.len()).step_by(2) {
        let condition = match &args[i] {
            CellValue::Boolean(b) => *b,
            CellValue::Number(n) => *n != 0.0,
            CellValue::Text(t) => !t.is_empty(),
            CellValue::Blank => false,
            CellValue::Formula(_) => return Err(EngineError::EvaluationError("Cannot use unevaluated formula in IFS".into())),
            CellValue::Error(e) => return Err(EngineError::CellValueError(e.clone())),
        };
        
        if condition {
            return Ok(args[i + 1].clone());
        }
    }
    
    // No conditions were true
    Err(EngineError::EvaluationError("No TRUE conditions in IFS function".into()))
}

// ===== LOOKUP & REFERENCE FUNCTIONS =====

// XLOOKUP function - flexible modern lookup (replaces VLOOKUP/HLOOKUP)
fn xlookup(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 6 {
        return Err(EngineError::EvaluationError(
            "XLOOKUP requires 3-6 arguments: lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]".into()));
    }
    
    // In a real implementation, we would need to handle arrays
    // This is a simplified version that only handles the current arguments
    
    // Default value if not found
    let if_not_found = if args.len() >= 4 { args[3].clone() } else { CellValue::Error("#N/A".to_string()) };
    
    // Placeholder implementation - in a real implementation, we would search through lookup_array for lookup_value
    // and return the corresponding value from return_array
    
    // For now, just returning a placeholder result
    // In a real implementation, we would check if we found a match and return if_not_found if necessary
    Ok(CellValue::Text("XLOOKUP Result".to_string()))
}

// XMATCH function - returns position of lookup_value in an array
fn xmatch(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 4 {
        return Err(EngineError::EvaluationError(
            "XMATCH requires 2-4 arguments: lookup_value, lookup_array, [match_mode], [search_mode]".into()));
    }
    
    // In a real implementation, we would need to handle arrays
    // This is a simplified version that only handles the current arguments
    
    // Placeholder implementation - in a real implementation, we would search through lookup_array for lookup_value
    // and return its position
    
    // For now, just returning a placeholder result
    Ok(CellValue::Number(3.0)) // Mock result - position 3
}

// OFFSET function - returns a range shifted from reference by given rows/cols
fn offset(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 3 || args.len() > 5 {
        return Err(EngineError::EvaluationError(
            "OFFSET requires 3-5 arguments: reference, rows, cols, [height], [width]".into()));
    }
    
    // Extract rows and columns to offset
    let rows = extract_number(&args[1], "rows")?;
    let cols = extract_number(&args[2], "cols")?;
    
    // Extract optional height and width
    let height = if args.len() >= 4 { extract_number(&args[3], "height")? } else { 1.0 };
    let width = if args.len() >= 5 { extract_number(&args[4], "width")? } else { 1.0 };
    
    if height <= 0.0 || width <= 0.0 {
        return Err(EngineError::EvaluationError("OFFSET height and width must be positive".into()));
    }
    
    // In a real implementation, we would need to handle cell references
    // This is a simplified version that only handles the current arguments
    
    // Placeholder implementation - in a real implementation, we would shift the reference
    // by the specified rows and cols and return a range with the specified height and width
    
    // For now, just returning a placeholder result
    Ok(CellValue::Number(42.0)) // Mock result
}

// INDIRECT function - interprets a text string as a cell or range reference
fn indirect(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 1 || args.len() > 2 {
        return Err(EngineError::EvaluationError(
            "INDIRECT requires 1-2 arguments: ref_text, [a1]".into()));
    }
    
    // Extract reference text
    let ref_text = match &args[0] {
        CellValue::Text(t) => t,
        _ => return Err(EngineError::EvaluationError("INDIRECT first argument must be text".into())),
    };
    
    // Extract A1 style flag (defaults to TRUE)
    let a1_style = if args.len() == 2 {
        match &args[1] {
            CellValue::Boolean(b) => *b,
            CellValue::Number(n) => *n != 0.0,
            _ => true,
        }
    } else {
        true
    };
    
    // In a real implementation, we would need to handle cell references
    // This is a simplified version that only handles the current arguments
    
    // Placeholder implementation - in a real implementation, we would interpret ref_text
    // as a cell reference and return its value
    
    // For now, just returning a placeholder result
    Ok(CellValue::Number(99.0)) // Mock result
}

// ===== DATE & TIME FUNCTIONS =====

// EOMONTH function - last day of month, offset by months
fn eomonth(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError(
            "EOMONTH requires exactly 2 arguments: start_date, months".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract months to add
    let months = extract_number(&args[1], "months")?;
    
    // In Excel, dates are stored as sequential serial numbers
    // 1 represents January 1, 1900
    // This is a simplified version that doesn't do the actual date calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual end of month date
    
    // For now, just returning a placeholder result
    Ok(CellValue::Number(start_date + 30.0 * months)) // Mock result
}

// EDATE function - same-day, offset by months
fn edate(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() != 2 {
        return Err(EngineError::EvaluationError(
            "EDATE requires exactly 2 arguments: start_date, months".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract months to add
    let months = extract_number(&args[1], "months")?;
    
    // In Excel, dates are stored as sequential serial numbers
    // 1 represents January 1, 1900
    // This is a simplified version that doesn't do the actual date calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual date
    
    // For now, just returning a placeholder result
    Ok(CellValue::Number(start_date + 30.0 * months)) // Mock result
}

// NETWORKDAYS function - count weekdays between dates
fn networkdays(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "NETWORKDAYS requires 2-3 arguments: start_date, end_date, [holidays]".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract end date
    let end_date = extract_number(&args[1], "end_date")?;
    
    if start_date > end_date {
        return Err(EngineError::EvaluationError("Start date must be less than or equal to end date".into()));
    }
    
    // In a real implementation, we would need to handle holidays and weekend days
    // This is a simplified version that doesn't do the actual calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual number of business days
    
    // For now, just returning a placeholder result
    // Roughly calculate workdays as 5/7 of the total days
    let total_days = end_date - start_date + 1.0;
    let approx_workdays = (total_days * 5.0 / 7.0).floor();
    
    Ok(CellValue::Number(approx_workdays))
}

// NETWORKDAYS.INTL function - customizable weekend
fn networkdays_intl(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 4 {
        return Err(EngineError::EvaluationError(
            "NETWORKDAYS.INTL requires 2-4 arguments: start_date, end_date, [weekend], [holidays]".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract end date
    let end_date = extract_number(&args[1], "end_date")?;
    
    if start_date > end_date {
        return Err(EngineError::EvaluationError("Start date must be less than or equal to end date".into()));
    }
    
    // Extract weekend parameter if provided
    let weekend_type = if args.len() >= 3 {
        match &args[2] {
            CellValue::Number(n) => *n,
            CellValue::Text(t) => {
                // In a real implementation, we would parse the weekend string here
                1.0 // Default to 1 (Saturday/Sunday) for now
            },
            _ => 1.0, // Default to 1 (Saturday/Sunday)
        }
    } else {
        1.0 // Default to 1 (Saturday/Sunday)
    };
    
    // In a real implementation, we would need to handle holidays and custom weekend days
    // This is a simplified version that doesn't do the actual calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual number of business days
    
    // For now, just returning a placeholder result similar to NETWORKDAYS
    // Roughly calculate workdays as 5/7 of the total days
    let total_days = end_date - start_date + 1.0;
    let approx_workdays = (total_days * 5.0 / 7.0).floor();
    
    Ok(CellValue::Number(approx_workdays))
}

// WORKDAY function - shift date by business days
fn workday(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "WORKDAY requires 2-3 arguments: start_date, days, [holidays]".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract days to add
    let days = extract_number(&args[1], "days")?;
    
    // In a real implementation, we would need to handle holidays and weekend days
    // This is a simplified version that doesn't do the actual calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual date
    
    // For now, just returning a placeholder result
    // Roughly calculate as adding days * 7/5 to account for weekends
    let approx_total_days = if days >= 0.0 {
        days * 7.0 / 5.0
    } else {
        days * 7.0 / 5.0
    };
    
    Ok(CellValue::Number(start_date + approx_total_days))
}

// WORKDAY.INTL function - customizable weekend
fn workday_intl(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 4 {
        return Err(EngineError::EvaluationError(
            "WORKDAY.INTL requires 2-4 arguments: start_date, days, [weekend], [holidays]".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract days to add
    let days = extract_number(&args[1], "days")?;
    
    // Extract weekend parameter if provided
    let weekend_type = if args.len() >= 3 {
        match &args[2] {
            CellValue::Number(n) => *n,
            CellValue::Text(t) => {
                // In a real implementation, we would parse the weekend string here
                1.0 // Default to 1 (Saturday/Sunday) for now
            },
            _ => 1.0, // Default to 1 (Saturday/Sunday)
        }
    } else {
        1.0 // Default to 1 (Saturday/Sunday)
    };
    
    // In a real implementation, we would need to handle holidays and custom weekend days
    // This is a simplified version that doesn't do the actual calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual date
    
    // For now, just returning a placeholder result similar to WORKDAY
    // Roughly calculate as adding days * 7/5 to account for weekends
    let approx_total_days = if days >= 0.0 {
        days * 7.0 / 5.0
    } else {
        days * 7.0 / 5.0
    };
    
    Ok(CellValue::Number(start_date + approx_total_days))
}

// YEARFRAC function - fraction of year between dates, per day-count basis
fn yearfrac(args: &[CellValue]) -> Result<CellValue, EngineError> {
    if args.len() < 2 || args.len() > 3 {
        return Err(EngineError::EvaluationError(
            "YEARFRAC requires 2-3 arguments: start_date, end_date, [basis]".into()));
    }
    
    // Extract start date
    let start_date = extract_number(&args[0], "start_date")?;
    // Extract end date
    let end_date = extract_number(&args[1], "end_date")?;
    
    // Extract basis (defaults to 0)
    let basis = if args.len() == 3 { extract_number(&args[2], "basis")? } else { 0.0 };
    
    if start_date < 0.0 || end_date < 0.0 || basis < 0.0 || basis > 4.0 || (basis != basis.floor()) {
        return Err(EngineError::EvaluationError("YEARFRAC arguments out of valid range".into()));
    }
    
    if start_date > end_date {
        return Err(EngineError::EvaluationError("Start date must be less than or equal to end date".into()));
    }
    
    // In a real implementation, we would need to handle different day count bases
    // This is a simplified version that doesn't do the actual calculation
    
    // Placeholder implementation - in a real implementation, we would calculate the actual fraction
    
    // For now, just returning a placeholder result
    // Simple calculation assuming 365-day year
    let days_between = end_date - start_date;
    let fraction = days_between / 365.0;
    
    Ok(CellValue::Number(fraction))
}
