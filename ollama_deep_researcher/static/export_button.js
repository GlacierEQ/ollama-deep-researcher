/**
 * Adds an export button to the LangGraph Studio UI
 * to export research results directly to Word
 */

(function () {
    function waitForElement(selector, callback, maxAttempts = 30, interval = 500) {
        let attempts = 0;

        const checkElement = () => {
            attempts++;
            const element = document.querySelector(selector);

            if (element) {
                callback(element);
                return;
            }

            if (attempts >= maxAttempts) {
                console.log(`Could not find element: ${selector} after ${maxAttempts} attempts`);
                return;
            }

            setTimeout(checkElement, interval);
        };

        checkElement();
    }

    function addExportButton() {
        waitForElement('.main-wrapper', (mainWrapper) => {
            // Check if our button already exists
            if (document.getElementById('export-to-word-button')) {
                return;
            }

            // Find the target container (toolbar)
            const toolbarContainer = mainWrapper.querySelector('.flex.flex-row.items-center.space-x-4');
            if (!toolbarContainer) {
                console.log('Could not find toolbar container');
                return;
            }

            // Create export button
            const exportButton = document.createElement('button');
            exportButton.id = 'export-to-word-button';
            exportButton.className = 'group-hover:bg-muted rounded-md px-2 py-1 hover:bg-white/10 text-sm flex items-center';
            exportButton.innerHTML = `
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" 
                     class="bi bi-file-earmark-word mr-1" viewBox="0 0 16 16">
                    <path d="M5.485 6.879a.5.5 0 1 0-.97.242l1.5 6a.5.5 0 0 0 .967.01L8 9.402l1.018 3.73a.5.5 0 0 0 .967-.01l1.5-6a.5.5 0 0 0-.97-.242l-1.036 4.144-.997-3.655a.5.5 0 0 0-.964 0l-.997 3.655L5.485 6.88z"/>
                    <path d="M14 14V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2zM9.5 3A1.5 1.5 0 0 0 11 4.5h2V14a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1h5.5v2z"/>
                </svg>
                Export to Word
            `;

            // Add click event
            exportButton.addEventListener('click', exportToWord);

            // Add to toolbar
            toolbarContainer.appendChild(exportButton);

            console.log('Export to Word button added');
        });
    }

    async function exportToWord() {
        try {
            // Show loading state
            const exportButton = document.getElementById('export-to-word-button');
            const originalContent = exportButton.innerHTML;
            exportButton.innerHTML = '<span class="animate-spin mr-1">↻</span> Exporting...';
            exportButton.disabled = true;

            // Fetch the current state
            const response = await fetch('/api/state');
            if (!response.ok) {
                throw new Error('Failed to fetch research data');
            }

            const data = await response.json();

            // Send to export endpoint
            const exportResponse = await fetch('/api/document/export-docx', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    markdown: data.final_summary,
                    topic: data.topic
                })
            });

            if (!exportResponse.ok) {
                throw new Error('Failed to export document');
            }

            // Download the file
            const blob = await exportResponse.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `research-${data.topic.substring(0, 20)}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            // Show success message
            exportButton.innerHTML = '<span style="color: #4CAF50">✓</span> Exported!';
            setTimeout(() => {
                exportButton.innerHTML = originalContent;
                exportButton.disabled = false;
            }, 2000);

        } catch (error) {
            console.error('Export failed:', error);
            alert(`Export failed: ${error.message}`);

            // Reset button
            const exportButton = document.getElementById('export-to-word-button');
            exportButton.innerHTML = '<span style="color: #F44336">✗</span> Export failed';
            setTimeout(() => {
                exportButton.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" 
                         class="bi bi-file-earmark-word mr-1" viewBox="0 0 16 16">
                        <path d="M5.485 6.879a.5.5 0 1 0-.97.242l1.5 6a.5.5 0 0 0 .967.01L8 9.402l1.018 3.73a.5.5 0 0 0 .967-.01l1.5-6a.5.5 0 0 0-.97-.242l-1.036 4.144-.997-3.655a.5.5 0 0 0-.964 0l-.997 3.655L5.485 6.88z"/>
                        <path d="M14 14V4.5L9.5 0H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2zM9.5 3A1.5 1.5 0 0 0 11 4.5h2V14a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1h5.5v2z"/>
                    </svg>
                    Export to Word
                `;
                exportButton.disabled = false;
            }, 2000);
        }
    }

    // Add the button when the page is loaded
    window.addEventListener('load', addExportButton);

    // Also add the button when the URL changes (SPA navigation)
    let lastUrl = location.href;
    new MutationObserver(() => {
        const url = location.href;
        if (url !== lastUrl) {
            lastUrl = url;
            setTimeout(addExportButton, 1000); // Wait for page to render
        }
    }).observe(document, { subtree: true, childList: true });
})();