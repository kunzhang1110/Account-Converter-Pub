python "./src/convert_nab.py"

if %errorlevel% neq 0 (
	echo An error has occurred. Press any key to continue...
	pause >nul
)

start  excel "./Converted.xlsx"

