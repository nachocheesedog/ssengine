// ssengine-core/src/ast.rs
// Abstract Syntax Tree for formula parsing

use crate::model::{CellAddress, RowId, ColumnId};

#[derive(Debug, Clone, PartialEq)]
pub enum AstNode {
    Literal(Literal),
    Reference(Reference),
    BinaryOp {
        op: BinaryOperator,
        left: Box<AstNode>,
        right: Box<AstNode>,
    },
    UnaryOp {
        op: UnaryOperator,
        operand: Box<AstNode>,
    },
    FunctionCall {
        name: String,
        args: Vec<AstNode>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub enum Literal {
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(crate::error::CellError),
}

#[derive(Debug, Clone, PartialEq)]
pub enum Reference {
    Cell(CellAddress),
    Range { start: CellAddress, end: CellAddress },
    SheetCell { sheet: String, address: CellAddress },
    SheetRange { sheet: String, start: CellAddress, end: CellAddress },
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
    Concat,
}

#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOperator {
    Positive,
    Negative,
    Percent,
}
