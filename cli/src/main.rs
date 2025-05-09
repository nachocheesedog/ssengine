// ssengine-cli/src/main.rs
// Command-line interface for ssengine

use clap::{Parser, Subcommand};
use ssengine_core::{Workbook, new_workbook};
use ssengine_io::{read_xlsx, write_xlsx};
use ssengine_sdk::run_server;
use std::path::{Path, PathBuf};
use std::net::{IpAddr, Ipv4Addr, SocketAddr};

#[derive(Parser)]
#[command(name = "ssengine")]
#[command(author, version, about, long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Create a new workbook
    New {
        /// Path to save the new workbook
        #[arg(short, long)]
        output: PathBuf,
    },
    
    /// Start the API server
    Serve {
        /// Port to listen on
        #[arg(short, long, default_value_t = 8080)]
        port: u16,
        
        /// Address to bind to
        #[arg(short, long, default_value = "127.0.0.1")]
        address: String,
        
        /// Optional workbook to load
        #[arg(short, long)]
        workbook: Option<PathBuf>,
    },
    
    /// Convert a workbook between formats
    Convert {
        /// Input file
        #[arg(short, long)]
        input: PathBuf,
        
        /// Output file
        #[arg(short, long)]
        output: PathBuf,
    },
}

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Initialize logging
    env_logger::init();
    
    // Parse command-line arguments
    let cli = Cli::parse();
    
    match cli.command {
        Commands::New { output } => {
            println!("Creating new workbook at {}", output.display());
            
            // Create a new workbook
            let wb = new_workbook();
            
            // Save it
            write_xlsx(&wb, output)?;
            
            println!("Workbook created successfully.");
        },
        
        Commands::Serve { port, address, workbook } => {
            println!("Starting ssengine API server on {}:{}", address, port);
            
            // Create workbook API
            let api = if let Some(path) = workbook {
                println!("Loading workbook from {}", path.display());
                let wb = read_xlsx(path)?;
                ssengine_sdk::WorkbookApi::from_workbook(wb)
            } else {
                println!("Creating empty workbook");
                ssengine_sdk::WorkbookApi::new()
            };
            
            // Parse the IP address
            let addr = address.parse::<IpAddr>().unwrap_or(IpAddr::V4(Ipv4Addr::new(127, 0, 0, 1)));
            let socket = SocketAddr::new(addr, port);
            
            // Run the server
            run_server(api, socket).await;
        },
        
        Commands::Convert { input, output } => {
            println!("Converting {} to {}", input.display(), output.display());
            
            // Infer file types from extensions
            let input_ext = input.extension().unwrap_or_default().to_string_lossy().to_lowercase();
            let output_ext = output.extension().unwrap_or_default().to_string_lossy().to_lowercase();
            
            // Load the input file
            let wb = if input_ext == "xlsx" {
                read_xlsx(input)?
            } else {
                return Err(format!("Unsupported input format: {}", input_ext).into());
            };
            
            // Save to the output format
            if output_ext == "xlsx" {
                write_xlsx(&wb, output)?
            } else {
                return Err(format!("Unsupported output format: {}", output_ext).into());
            }
            
            println!("Conversion completed successfully.");
        },
    }
    
    Ok(())
}
