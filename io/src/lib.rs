// ssengine-io/src/lib.rs

pub mod xlsx;
pub mod csv;

// Re-export key functionality
pub use xlsx::{read_xlsx, write_xlsx};
pub use csv::{read_csv, write_csv};

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
