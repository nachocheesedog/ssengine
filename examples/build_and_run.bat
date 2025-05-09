@echo off
echo Building ssengine workspace...
cd ..
cargo build

echo.
echo Running simple_model example...
cargo run --example simple_model

echo.
echo Done!
pause
