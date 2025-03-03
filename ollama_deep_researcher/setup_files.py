import os
import base64
import shutil
from pathlib import Path

def setup_word_addin():
    """Set up all necessary files for Word add-in integration"""
    print("Setting up Word Add-in files...")
    
    # Define paths
    base_dir = Path(__file__).parent
    static_dir = base_dir / "static"
    static_dir.mkdir(exist_ok=True)
    
    # Create icon files
    create_icons(static_dir)
    
    # Copy manifest file
    create_manifest_file(static_dir)
    
    # Create installation batch file
    create_install_batch(static_dir)
    
    # Create Word add-in HTML interface
    create_addin_html(static_dir)
    
    print("Word Add-in files setup complete!")

def create_icons(static_dir):
    """Create icon files for the add-in"""
    print("Creating icon files...")
    
    # Simple base64-encoded icon for the add-in (blue square with "OR" text)
    icon32_b64 = """
    iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABK0lEQVR42mNkGIqAEd3ntbW1/
    9EwVqW1tbWMo3YMWQfU1dX9pyYeDAEYBQ7A5X9wayG7npKUAHfAUHIAC7q/8YU+LrigdmFEsQ
    TOEBhFMOAAXFEQGRkJZmdnZ4MVbGxswl2Ym5sLVlRQUABm+/v7g9mXLl0apA5gRg//jo4OMBt
    W8iEDWLTAQh4XGEDLUKKgoAAliYH4e/fuQZwwcA5AzgWwdI/NAfDEhy8aDAwMBq8DmJFzAXLa
    R04DyLkAvRhGLsjHGwVoDsCXBoaj/wctCpBdAEv7sJIQW96AVT2jDqBuWTCKYPAyIbIjkNM5c
    skHw0OpTzB0HcBAQNPAOmAUDYYoYCLFACMjI1w5rK06WN4IwQAWMBNGQ2BwRAG1AWPbkAIAAL
    wF27iJXhCHAAAAAElFTkSuQmCC
    """

    icon80_b64 = """
    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
    gg==
    """

    # Create icon files
    icon32_path = static_dir / "icon-32.png"
    icon80_path = static_dir / "icon-80.png"

    try:
        if not icon32_path.exists():
            with open(icon32_path, 'wb') as f:
                f.write(base64.b64decode(icon32_b64.strip()))
            print(f"Created {icon32_path}")
        
        if not icon80_path.exists():
            with open(icon80_path, 'wb') as f:
                f.write(base64.b64decode(icon80_b64.strip()))
            print(f"Created {icon80_path}")
    except Exception as e:
        print(f"Warning: Could not create icon files: {str(e)}")

def create_manifest_file(static_dir):
    """Create Word add-in manifest file"""
    print("Creating manifest file...")
    
    manifest_path = static_dir / "manifest.xml"
    
    manifest_content = """<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
          xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
          xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
          xsi:type="TaskPaneApp">

  <!-- Basic Settings: Add-in metadata, used for all versions of Office unless override provided. -->
  <Id>37b5b3d9-c7ce-4106-9a3b-af7f0c0c7e4c</Id>
  <Version>1.0.0</Version>
  <ProviderName>Ollama Deep Researcher</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Ollama Deep Researcher" />
  <Description DefaultValue="Deep research and content generation powered by Ollama" />
  <IconUrl DefaultValue="http://localhost:2024/static/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="http://localhost:2024/static/icon-80.png" />
  <SupportUrl DefaultValue="https://github.com/langchain-ai/ollama-deep-researcher" />

  <!-- Domains that will be allowed when navigating. -->
  <AppDomains>
    <AppDomain>localhost</AppDomain>
  </AppDomains>

  <!-- Hosts supported by this add-in -->
  <Hosts>
    <Host Name="Document" />
  </Hosts>

  <!-- Default settings -->
  <DefaultSettings>
    <SourceLocation DefaultValue="http://localhost:2024/word-addin" />
  </DefaultSettings>

  <!-- Permissions needed -->
  <Permissions>ReadWriteDocument</Permissions>

  <!-- Version Overrides -->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
          
          <!-- Ribbon Integration -->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label" />
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16" />
                  <bt:Image size="32" resid="Icon.32x32" />
                  <bt:Image size="80" resid="Icon.80x80" />
                </Icon>
                
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    
    <!-- Resources -->
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="http://localhost:2024/static/icon-32.png" />
        <bt:Image id="Icon.32x32" DefaultValue="http://localhost:2024/static/icon-32.png" />
        <bt:Image id="Icon.80x80" DefaultValue="http://localhost:2024/static/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="http://localhost:2024/word-addin" />
        <bt:Url id="Taskpane.Url" DefaultValue="http://localhost:2024/word-addin" />
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/langchain-ai/ollama-deep-researcher" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with the Ollama Deep Researcher!" />
        <bt:String id="CommandsGroup.Label" DefaultValue="Ollama Tools" />
        <bt:String id="TaskpaneButton.Label" DefaultValue="Ollama Deep Researcher" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="The Ollama Deep Researcher add-in is now loaded. Go to the Home tab and click the 'Ollama Deep Researcher' button to get started." />
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to open the Ollama Deep Researcher for research and content help." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""
    
    with open(manifest_path, 'w') as f:
        f.write(manifest_content)
    
    print(f"Created {manifest_path}")

def create_install_batch(static_dir):
    """Create installation batch script"""
    print("Creating installation script...")
    
    batch_path = static_dir / "install-word-addin.bat"
    
    batch_content = """@echo off
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
"""
    
    with open(batch_path, 'w') as f:
        f.write(batch_content)
    
    print(f"Created {batch_path}")

def create_addin_html(static_dir):
    """Create Word add-in HTML interface"""
    print("Creating Word add-in HTML interface...")
    
    html_path = static_dir / "word-addin.html"
    
    # HTML content (abbreviated for clarity)
    html_content = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Ollama Deep Researcher</title>
    <!-- Office.js CDN -->
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

    <style>
        body {
            font-family: 'Segoe UI', sans-serif;
            margin: 0;
            padding: 15px;
        }
        #app {
            display: flex;
            flex-direction: column;
            height: 100vh;
        }
        .header {
            padding: 10px 0;
            border-bottom: 1px solid #ddd;
            margin-bottom: 15px;
        }
        .logo {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 10px;
        }
        .logo img {
            width: 32px;
            height: 32px;
        }
        .logo h1 {
            font-size: 16px;
            margin: 0;
        }
        .tabs {
            display: flex;
            border-bottom: 1px solid #ddd;
            margin-bottom: 15px;
        }
        .tab {
            padding: 8px 15px;
            cursor: pointer;
            background: #f5f5f5;
            border: 1px solid #ddd;
            border-bottom: none;
            margin-right: 5px;
            border-radius: 5px 5px 0 0;
        }
        .tab.active {
            background: #fff;
            border-bottom: 1px solid white;
            margin-bottom: -1px;
        }
        .tab-content {
            display: none;
            flex-grow: 1;
            overflow-y: auto;
        }
        .tab-content.active {
            display: block;
        }
        .form-group {
            margin-bottom: 15px;
        }
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        textarea, input[type="text"] {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
        }
        textarea {
            height: 150px;
            resize: vertical;
        }
        button {
            background-color: #0066cc;
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: bold;
        }
        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
        .result {
            margin-top: 20px;
            border-top: 1px solid #ddd;
            padding-top: 15px;
        }
        .result pre {
            white-space: pre-wrap;
            background-color: #f9f9f9;
            padding: 10px;
            border-radius: 4px;
            overflow-x: auto;
        }
        .spinner {
            border: 4px solid rgba(0, 0, 0, 0.1);
            width: 24px;
            height: 24px;
            border-radius: 50%;
            border-left-color: #0066cc;
            animation: spin 1s linear infinite;
            display: inline-block;
            vertical-align: middle;
            margin-right: 10px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .status {
            margin-top: 10px;
            font-style: italic;
        }
        .error {
            color: #cc0000;
            margin-top: 10px;
        }
        .button-row {
            display: flex;
            justify-content: space-between;
            margin-top: 15px;
        }
        .markdown {
            line-height: 1.6;
        }
        .markdown h1, .markdown h2 {
            border-bottom: 1px solid #eee;
            padding-bottom: 8px;
        }
    </style>
</head>
<body>
    <div id="app">
        <div class="header">
            <div class="logo">
                <img src="/static/icon-32.png" alt="Ollama Deep Researcher">
                <h1>Ollama Deep Researcher</h1>
            </div>
            <p>Powered by your local Ollama LLM</p>
        </div>

        <div class="tabs">
            <div class="tab active" data-tab="edit-tab">Edit</div>
            <div class="tab" data-tab="research-tab">Research</div>
        </div>

        <div class="tab-content active" id="edit-tab">
            <div class="form-group">
                <label for="edit-content">Selected Text</label>
                <textarea id="edit-content" rows="6" placeholder="Select text in your document first, then click 'Get Selected Text'"></textarea>
            </div>
            <div class="form-group">
                <label for="edit-instruction">Editing Instruction</label>
                <input type="text" id="edit-instruction" placeholder="e.g., Make it more concise or Change the tone to be more professional">
            </div>
            <div class="button-row">
                <button id="get-selected-text">Get Selected Text</button>
                <button id="edit-button">Edit</button>
            </div>
            <div class="status" id="edit-status"></div>
            <div class="error" id="edit-error"></div>
            <div class="result" id="edit-result" style="display: none;">
                <h3>Edited Text</h3>
                <div id="edit-output"></div>
                <button id="insert-edit" style="margin-top: 10px;">Insert into Document</button>
            </div>
        </div>

        <div class="tab-content" id="research-tab">
            <div class="form-group">
                <label for="research-topic">Research Topic</label>
                <input type="text" id="research-topic" placeholder="Enter a topic to research">
            </div>
            <div class="form-group">
                <label for="research-context">Context (Optional)</label>
                <textarea id="research-context" rows="4" placeholder="Add any additional context or requirements"></textarea>
            </div>
            <button id="research-button">Research</button>
            <div class="status" id="research-status"></div>
            <div class="error" id="research-error"></div>
            <div class="result" id="research-result" style="display: none;">
                <h3>Research Results</h3>
                <div id="research-output" class="markdown"></div>
                <button id="insert-research" style="margin-top: 10px;">Insert into Document</button>
            </div>
        </div>
    </div>

    <script>
        // Initialize Office.js
        Office.onReady(function(info) {
            if (info.host === Office.HostType.Word) {
                initializeApp();
            } else {
                document.body.innerHTML = "<p>This add-in only works in Microsoft Word.</p>";
            }
        });

        function initializeApp() {
            // Tab navigation
            document.querySelectorAll('.tab').forEach(tab => {
                tab.addEventListener('click', () => {
                    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                    tab.classList.add('active');
                    document.getElementById(tab.dataset.tab).classList.add('active');
                });
            });

            // Get selected text button
            document.getElementById('get-selected-text').addEventListener('click', getSelectedText);
            
            // Edit button
            document.getElementById('edit-button').addEventListener('click', editContent);
            
            // Insert edited text button
            document.getElementById('insert-edit').addEventListener('click', insertEditedText);
            
            // Research button
            document.getElementById('research-button').addEventListener('click', researchTopic);
            
            // Insert research button
            document.getElementById('insert-research').addEventListener('click', insertResearchText);
        }

        function getSelectedText() {
            const editStatus = document.getElementById('edit-status');
            editStatus.textContent = "Getting selected text...";
            
            Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                
                await context.sync();
                
                document.getElementById('edit-content').value = selection.text;
                editStatus.textContent = "Text loaded from document.";
                
                setTimeout(() => {
                    editStatus.textContent = "";
                }, 3000);
            })
            .catch(error => {
                document.getElementById('edit-error').textContent = `Error: ${error.message}`;
                console.error(error);
            });
        }

        async function editContent() {
            // Implementation for edit functionality
            // ...
        }

        function insertEditedText() {
            // Implementation for inserting edited text
            // ...
        }

        async function researchTopic() {
            // Implementation for research functionality
            // ...
        }

        function insertResearchText() {
            // Implementation for inserting research
            // ...
        }

        // Simple markdown to HTML converter
        function markdownToHtml(markdown) {
            // Implementation of markdown converter
            // ...
        }
    </script>
</body>
</html>
"""
    
    with open(html_path, 'w') as f:
        f.write(html_content)
    
    print(f"Created {html_path}")

# Run setup if this script is executed directly
if __name__ == "__main__":
    setup_word_addin()