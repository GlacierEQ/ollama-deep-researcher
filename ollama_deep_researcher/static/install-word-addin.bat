@echo off
echo Installing Ollama Deep Researcher Word Add-in...
set MANIFEST_URL=http://localhost:2024/static/manifest.xml
set STARTUP_DIR=%APPDATA%\Microsoft\Word\STARTUP
set MANIFEST_PATH=%STARTUP_DIR%\ollama-researcher-manifest.xml

mkdir "%STARTUP_DIR%" 2>nul

echo Downloading manifest...
powershell -Command "Invoke-WebRequest -Uri '%MANIFEST_URL%' -OutFile '%MANIFEST_PATH%'"

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to download manifest. Make sure the server is running at http://localhost:2024
    pause
    exit /b 1
)

echo Add-in installed successfully!
echo.
echo Please restart Microsoft Word and look for "Ollama Deep Researcher" in the Add-ins section.
echo.
pause

## Troubleshooting

If the add-in doesn't appear:
1. Make sure the local server is running (`langgraph dev`)
2. Check that your browser can access https://localhost:2024/word-addin
3. Try restarting Word after installation

## Using the Add-in

1. Open Word and go to the **Home** tab
2. Click on the **Ollama Deep Researcher** button in the ribbon
3. In the sidebar that opens:
    - **Edit**: Select text, enter editing instructions, and click "Edit" to refine your content
    - **Research**: Enter a topic and click "Research" to generate content
