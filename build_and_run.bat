@echo off
echo Checking and building ssengine workspace...

echo.
echo Step 1: Check the workspace for errors
cargo check

IF %ERRORLEVEL% NEQ 0 (
  echo Workspace check failed, see errors above.
  pause
  exit /b %ERRORLEVEL%
)

echo.
echo Step 2: Building minimal example...
cargo build --bin minimal

IF %ERRORLEVEL% NEQ 0 (
  echo Build failed, see errors above.
  pause
  exit /b %ERRORLEVEL%
)

echo.
echo Step 3: Running minimal example...
cargo run --bin minimal

echo.
echo Done! An XLSX file has been created.
pause
