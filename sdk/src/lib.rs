// ssengine-sdk/src/lib.rs

pub mod api;
pub mod server;
pub mod schemas;

// Re-export key functionality for easier access
pub use server::run_server;
pub use api::WorkbookApi;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
