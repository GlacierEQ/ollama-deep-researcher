/**
 * Ollama Deep Researcher - Word Add-in functionality
 */

// Global state
let apiEndpoint = '/api';
let maxDocumentSize = 20000; // Characters
let isProcessing = false;

// Initialize Office.js
Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        initializeApp();
        checkApiStatus();
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

    // Get selected text buttons
    document.getElementById('get-selected-text').addEventListener('click', () => getSelectedText('edit'));
    document.getElementById('get-verify-text').addEventListener('click', () => getSelectedText('verify'));

    // Edit button
    document.getElementById('edit-button').addEventListener('click', editContent);

    // Insert edited text button
    document.getElementById('insert-edit').addEventListener('click', insertEditedText);

    // Research button
    document.getElementById('research-button').addEventListener('click', researchTopic);

    // Insert research button
    document.getElementById('insert-research').addEventListener('click', insertResearchText);

    // Verify button
    document.getElementById('verify-button').addEventListener('click', verifyContent);
}

async function checkApiStatus() {
    try {
        const response = await fetch(`${apiEndpoint}/word/status`);
        if (response.ok) {
            const data = await response.json();
            if (data.model_responsive) {
                console.log("API status check: model is responsive.");
            } else {
                showWarning("edit", "The Ollama model may not be available. Some features might not work correctly.");
            }
        } else {
            showWarning("edit", "Unable to connect to the API. Please check if the server is running.");
        }
    } catch (error) {
        console.error("API status check failed:", error);
        showWarning("edit", "Unable to connect to the API. Please check if the server is running.");
    }
}

function getSelectedText(target) {
    if (isProcessing) {
        showWarning(target, "Another operation is in progress. Please wait...");
        return;
    }

    const statusElement = document.getElementById(`${target}-status`);
    const errorElement = document.getElementById(`${target}-error`);
    const warningElement = document.getElementById(`${target}-warning`);
    const contentElement = document.getElementById(`${target}-content`);

    statusElement.textContent = "Getting selected text...";
    errorElement.textContent = "";
    warningElement.textContent = "";

    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");

        await context.sync();

        const selectedText = selection.text;
        if (!selectedText || selectedText.trim().length === 0) {
            errorElement.textContent = "No text selected. Please select some text in your document.";
            statusElement.textContent = "";
            return;
        }

        contentElement.value = selectedText;
        statusElement.textContent = "Text loaded from document.";

        // Check if text is very long
        if (selectedText.length > 8000) {
            showWarning(target, `Selected text is quite long (${selectedText.length} characters). Processing may take longer.`);
        }

        if (selectedText.length > maxDocumentSize) {
            errorElement.textContent = `The selected text is too long (${selectedText.length} characters). Maximum allowed is ${maxDocumentSize}.`;
            warningElement.textContent = "Consider selecting a smaller portion of text.";
            return;
        }

        setTimeout(() => {
            statusElement.textContent = "";
        }, 3000);
    })
        .catch(error => {
            errorElement.textContent = `Error: ${error.message}`;
            console.error(error);
        });
}

async function editContent() {
    const content = document.getElementById('edit-content').value;
    const instruction = document.getElementById('edit-instruction').value;

    if (!content) {
        document.getElementById('edit-error').textContent = "Please select or enter text to edit.";
        return;
    }

    if (!instruction) {
        document.getElementById('edit-error').textContent = "Please enter an editing instruction.";
        return;
    }

    if (isProcessing) {
        document.getElementById('edit-warning').textContent = "Another operation is in progress. Please wait...";
        return;
    }

    isProcessing = true;
    clearMessages("edit");
    document.getElementById('edit-status').textContent = "Editing text...";
    document.getElementById('edit-button').disabled = true;

    try {
        const response = await fetch(`${apiEndpoint}/word/edit`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ content, instruction })
        });

        const data = await response.json();

        if (response.ok) {
            document.getElementById('edit-output').textContent = data.edited_content;
            document.getElementById('edit-result').style.display = "block";

            if (data.processing_time_seconds) {
                document.getElementById('edit-status').textContent =
                    `Editing complete in ${data.processing_time_seconds} seconds.`;
            } else {
                document.getElementById('edit-status').textContent = "Editing complete.";
            }

            if (data.warning) {
                document.getElementById('edit-warning').textContent = data.warning;
            }

            if (data.info) {
                document.getElementById('edit-info').textContent = data.info;
            }

            // Auto-scroll to result
            document.getElementById('edit-result').scrollIntoView({ behavior: 'smooth' });
        } else {
            document.getElementById('edit-error').textContent = data.error || "Failed to edit content.";
        }
    } catch (error) {
        document.getElementById('edit-error').textContent = "Error: " + error.message;
    } finally {
        document.getElementById('edit-button').disabled = false;
        isProcessing = false;
    }
}

function insertEditedText() {
    if (isProcessing) {
        showWarning("edit", "Another operation is in progress. Please wait...");
        return;
    }

    const editedText = document.getElementById('edit-output').textContent;

    if (!editedText) {
        document.getElementById('edit-error').textContent = "No edited text to insert.";
        return;
    }

    isProcessing = true;
    document.getElementById('edit-status').textContent = "Inserting text...";

    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(editedText, Word.InsertLocation.replace);

        await context.sync();

        document.getElementById('edit-status').textContent = "Text inserted into document.";

        setTimeout(() => {
            document.getElementById('edit-status').textContent = "";
        }, 3000);
    })
        .catch(error => {
            document.getElementById('edit-error').textContent = `Error: ${error.message}`;
        })
        .finally(() => {
            isProcessing = false;
        });
}

async function researchTopic() {
    const topic = document.getElementById('research-topic').value;
    const context = document.getElementById('research-context').value;

    if (!topic) {
        document.getElementById('research-error').textContent = "Please enter a research topic.";
        return;
    }

    if (isProcessing) {
        document.getElementById('research-warning').textContent = "Another operation is in progress. Please wait...";
        return;
    }

    isProcessing = true;
    clearMessages("research");
    document.getElementById('research-status').textContent = "Researching topic... This may take a minute or two.";
    document.getElementById('research-button').disabled = true;
    document.getElementById('research-result').style.display = "none";

    // Show spinner
    const spinner = document.createElement('div');
    spinner.className = 'spinner';
    spinner.id = 'research-spinner';
    document.getElementById('research-status').prepend(spinner);

    try {
        const response = await fetch(`${apiEndpoint}/word/research`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ topic, context })
        });

        const data = await response.json();

        // Remove spinner
        const spinnerElement = document.getElementById('research-spinner');
        if (spinnerElement) spinnerElement.remove();

        if (response.ok) {
            const markdownContent = data.markdown_content;
            document.getElementById('research-output').innerHTML = markdownToHtml(markdownContent);
            document.getElementById('research-result').style.display = "block";
            document.getElementById('research-status').textContent = "Research complete.";
            // Store the markdown for insertion
            document.getElementById('research-output').dataset.markdown = markdownContent;

            // Update metrics
            document.getElementById('research-time').textContent = data.processing_time_seconds || "-";
            document.getElementById('research-sources').textContent = data.source_count || "-";

            // Auto-scroll to result
            document.getElementById('research-result').scrollIntoView({ behavior: 'smooth' });
        } else {
            document.getElementById('research-error').textContent = data.error || "Failed to research topic.";
        }
    } catch (error) {
        document.getElementById('research-error').textContent = "Error: " + error.message;
    } finally {
        document.getElementById('research-button').disabled = false;
        isProcessing = false;
    }
}

function insertResearchText() {
    if (isProcessing) {
        showWarning("research", "Another operation is in progress. Please wait...");
        return;
    }

    const markdownContent = document.getElementById('research-output').dataset.markdown;

    if (!markdownContent) {
        document.getElementById('research-error').textContent = "No research content to insert.";
        return;
    }

    isProcessing = true;
    document.getElementById('research-status').textContent = "Inserting research...";

    Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(markdownContent, Word.InsertLocation.replace);

        await context.sync();

        document.getElementById('research-status').textContent = "Research inserted into document.";

        setTimeout(() => {
            document.getElementById('research-status').textContent = "";
        }, 3000);
    })
        .catch(error => {
            document.getElementById('research-error').textContent = `Error: ${error.message}`;
        })
        .finally(() => {
            isProcessing = false;
        });
}

async function verifyContent() {
    const content = document.getElementById('verify-content').value;

    if (!content) {
        document.getElementById('verify-error').textContent = "Please select or enter text to verify.";
        return;
    }

    if (isProcessing) {
        document.getElementById('verify-warning').textContent = "Another operation is in progress. Please wait...";
        return;
    }

    isProcessing = true;
    clearMessages("verify");
    document.getElementById('verify-status').textContent = "Analyzing writing...";
    document.getElementById('verify-button').disabled = true;

    // Show spinner
    const spinner = document.createElement('div');
    spinner.className = 'spinner';
    spinner.id = 'verify-spinner';
    document.getElementById('verify-status').prepend(spinner);

    try {
        const response = await fetch(`${apiEndpoint}/word/verify`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ content })
        });

        const data = await response.json();

        // Remove spinner
        const spinnerElement = document.getElementById('verify-spinner');
        if (spinnerElement) spinnerElement.remove();

        if (response.ok) {
            document.getElementById('verify-result').style.display = "block";

            // Update metrics
            document.getElementById('total-issues').textContent = data.total_issues || "0";
            document.getElementById('grammar-issues').textContent = (data.grammar_issues?.length || 0);
            document.getElementById('style-issues').textContent = (data.style_issues?.length || 0);
            document.getElementById('structure-issues').textContent = (data.structure_issues?.length || 0);

            // Display overall assessment
            document.getElementById('overall-assessment').textContent = data.overall_assessment || "No assessment provided.";

            // Populate issue lists
            populateIssueList('grammar-list', data.grammar_issues);
            populateIssueList('style-list', data.style_issues);
            populateIssueList('structure-list', data.structure_issues);
            populateIssueList('suggestions-list', data.improvement_suggestions);

            document.getElementById('verify-status').textContent = "Analysis complete.";

            // Auto-scroll to result
            document.getElementById('verify-result').scrollIntoView({ behavior: 'smooth' });
        } else {
            document.getElementById('verify-error').textContent = data.error || "Failed to analyze content.";
        }
    } catch (error) {
        document.getElementById('verify-error').textContent = "Error: " + error.message;
    } finally {
        document.getElementById('verify-button').disabled = false;
        isProcessing = false;
    }
}

function populateIssueList(listId, issues) {
    const listElement = document.getElementById(listId);
    listElement.innerHTML = "";

    if (!issues || issues.length === 0) {
        const listItem = document.createElement('li');
        listItem.className = 'issue-item';
        listItem.textContent = "No issues found.";
        listElement.appendChild(listItem);
        return;
    }

    issues.forEach(issue => {
        const listItem = document.createElement('li');
        listItem.className = 'issue-item';
        listItem.textContent = issue;
        listElement.appendChild(listItem);
    });
}

function clearMessages(target) {
    document.getElementById(`${target}-info`).textContent = "";
    document.getElementById(`${target}-status`).textContent = "";
    document.getElementById(`${target}-warning`).textContent = "";
    document.getElementById(`${target}-error`).textContent = "";

    // Remove any spinners
    const spinner = document.getElementById(`${target}-spinner`);
    if (spinner) spinner.remove();
}

function showWarning(target, message) {
    document.getElementById(`${target}-warning`).textContent = message;
}

// Simple markdown to HTML converter
function markdownToHtml(markdown) {
    if (!markdown) return '';

    let html = markdown
        // Headers
        .replace(/^### (.*$)/gm, '<h3>$1</h3>')
        .replace(/^## (.*$)/gm, '<h2>$1</h2>')
        .replace(/^# (.*$)/gm, '<h1>$1</h1>')

        // Lists
        .replace(/^\* (.*$)/gm, '<ul><li>$1</li></ul>')
        .replace(/^- (.*$)/gm, '<ul><li>$1</li></ul>')
        .replace(/^\d\. (.*$)/gm, '<ol><li>$1</li></ol>')

        // Bold and italic
        .replace(/\*\*(.*?)\*\*/gm, '<strong>$1</strong>')
        .replace(/\*(.*?)\*/gm, '<em>$1</em>')

        // Links
        .replace(/\[(.*?)\]\((.*?)\)/gm, '<a href="$2" target="_blank">$1</a>')

        // Line breaks
        .replace(/\n$/gm, '<br />');

    // Fix duplicate list tags
    html = html
        .replace(/<\/ul><ul>/g, '')
        .replace(/<\/ol><ol>/g, '');

    // Convert paragraphs
    const paragraphs = html.split(/\n\n+/);
    html = paragraphs.map(p => {
        if (p.indexOf('<h') === 0 || p.indexOf('<ul') === 0 || p.indexOf('<ol') === 0) {
            return p;
        }
        return `<p>${p}</p>`;
    }).join('');

    return html;
}
