@echo off
echo Building ssengine minimal example...
cd ..
cargo build --bin minimal

IF %ERRORLEVEL% NEQ 0 (
  echo Build failed, see errors above.
  pause
  exit /b %ERRORLEVEL%
)

echo.
echo Running minimal example...
cargo run --bin minimal

echo.
echo Done!
pause
