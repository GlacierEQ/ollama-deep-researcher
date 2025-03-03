(function () {
    "use strict";

    // Global variables
    let apiBaseUrl = "http://localhost:2024";
    let currentResearchResults = null;
    let isResearching = false;

    // The Office initialize function must be run each time a new page is loaded
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Word) {
            // Initialize UI elements
            document.getElementById("researchButton").onclick = doResearch;
            document.getElementById("insertButton").onclick = insertResearch;
            document.getElementById("improveButton").onclick = improveSelection;
            document.getElementById("expandButton").onclick = expandSelection;
            
            // Load settings from localStorage if available
            if (localStorage.getItem("ollamaModel")) {
                document.getElementById("ollamaModel").value = localStorage.getItem("ollamaModel");
            }
            if (localStorage.getItem("researchDepth")) {
                document.getElementById("researchDepth").value = localStorage.getItem("researchDepth");
            }
            
            // Save settings when changed
            document.getElementById("ollamaModel").onchange = saveSettings;
            document.getElementById("researchDepth").onchange = saveSettings;
        }
    });

    // Save settings to localStorage
    function saveSettings() {
        localStorage.setItem("ollamaModel", document.getElementById("ollamaModel").value);
        localStorage.setItem("researchDepth", document.getElementById("researchDepth").value);
    }

    // Perform research using the API
    async function doResearch() {
        const topic = document.getElementById("researchTopic").value.trim();
        if (!topic) {
            updateStatus("Please enter a research topic", "error");
            return;
        }

        // Get settings
        const model = document.getElementById("ollamaModel").value;
        const depth = document.getElementById("researchDepth").value;
        
        try {
            isResearching = true;
            updateStatus("Researching topic: " + topic + "...");
            document.getElementById("researchButton").disabled = true;
            
            // Call the research API
            const response = await fetch(`${apiBaseUrl}/api/research`, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    topic: topic,
                    model: model,
                    max_loops: parseInt(depth, 10)
                })
            });
            
            if (!response.ok) {
                throw new Error(`API error: ${response.status}`);
            }
            
            // Get the thread ID from the response
            const data = await response.json();
            const threadId = data.thread_id;
            
            // Poll for results
            await pollForResults(threadId);
            
        } catch (error) {
            console.error("Research error:", error);
            updateStatus("Research failed: " + error.message, "error");
        } finally {
            document.getElementById("researchButton").disabled = false;
            isResearching = false;
        }
    }
    
    // Poll for research results
    async function pollForResults(threadId) {
        try {
            // Wait a moment before first poll
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            let completed = false;
            let attempts = 0;
            const maxAttempts = 120; // 2 minutes max
            
            while (!completed && attempts < maxAttempts) {
                attempts++;
                
                updateStatus(`Researching... (${attempts}s)`);
                
                const response = await fetch(`${apiBaseUrl}/api/research/${threadId}/status`);
                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }
                
                const data = await response.json();
                if (data.status === "completed") {
                    completed = true;
                    // Get the final results
                    const resultsResponse = await fetch(`${apiBaseUrl}/api/research/${threadId}/results`);
                    if (!resultsResponse.ok) {
                        throw new Error(`API error: ${resultsResponse.status}`);
                    }
                    
                    currentResearchResults = await resultsResponse.json();
                    updateStatus("Research completed! You can now insert the results into your document.");
                    document.getElementById("insertButton").disabled = false;
                    
                } else if (data.status === "failed") {
                    throw new Error("Research process failed");
                } else {
                    // Wait 1 second before polling again
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }
            }
            
            if (!completed) {
                throw new Error("Research timed out");
            }
            
        } catch (error) {
            console.error("Polling error:", error);
            updateStatus("Research failed: " + error.message, "error");
        }
    }
    
    // Insert research results into the document
    async function insertResearch() {
        if (!currentResearchResults || !currentResearchResults.final_summary) {
            updateStatus("No research results available", "error");
            return;
        }
        
        try {
            // Insert the final summary at the current selection
            await Word.run(async (context) => {
                const markdown = currentResearchResults.final_summary;
                
                // Convert markdown to Word-friendly format
                // This is a simple implementation - a full markdown parser would be better
                let content = markdown;
                
                // Replace headers
                content = content.replace(/^# (.*?)$/gm, "$1\n======\n");
                content = content.replace(/^## (.*?)$/gm, "$1\n------\n");
                content = content.replace(/^### (.*?)$/gm, "$1\n");
                
                // Insert at current selection
                context.document.getSelection().insertText(content, Word.InsertLocation.replace);
                
                // Apply formatting (basic rich text)
                await context.sync();
                
                updateStatus("Research inserted into document!");
            });
        } catch (error) {
            console.error("Insert error:", error);
            updateStatus("Failed to insert research: " + error.message, "error");
        }
    }
    
    // Improve the selected text
    async function improveSelection() {
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                await context.sync();
                
                if (!selection.text || selection.text.trim() === "") {
                    updateStatus("Please select some text to improve", "error");
                    return;
                }
                
                updateStatus("Improving selected text...");
                document.getElementById("improveButton").disabled = true;
                
                // Call the improve API
                const model = document.getElementById("ollamaModel").value;
                const response = await fetch(`${apiBaseUrl}/api/improve`, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({
                        text: selection.text,
                        model: model
                    })
                });
                
                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }
                
                const data = await response.json();
                
                // Replace the selection with improved text
                selection.insertText(data.improved_text, Word.InsertLocation.replace);
                await context.sync();
                
                updateStatus("Text improved successfully!");
            });
        } catch (error) {
            console.error("Improve error:", error);
            updateStatus("Failed to improve text: " + error.message, "error");
        } finally {
            document.getElementById("improveButton").disabled = false;
        }
    }
    
    // Expand the selected text
    async function expandSelection() {
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load("text");
                await context.sync();
                
                if (!selection.text || selection.text.trim() === "") {
                    updateStatus("Please select some text to expand", "error");
                    return;
                }
                
                updateStatus("Expanding selected text...");
                document.getElementById("expandButton").disabled = true;
                
                // Call the expand API
                const model = document.getElementById("ollamaModel").value;
                const response = await fetch(`${apiBaseUrl}/api/expand`, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({
                        text: selection.text,
                        model: model
                    })
                });
                
                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }
                
                const data = await response.json();
                
                // Replace the selection with expanded text
                selection.insertText(data.expanded_text, Word.InsertLocation.replace);
                await context.sync();
                
                updateStatus("Text expanded successfully!");
            });
        } catch (error) {
            console.error("Expand error:", error);
            updateStatus("Failed to expand text: " + error.message, "error");
        } finally {
            document.getElementById("expandButton").disabled = false;
        }
    }
    
    // Update status message
    function updateStatus(message, type = "info") {
        const statusElement = document.getElementById("statusMessage");
        statusElement.textContent = message;
        statusElement.className = "status-box " + type;
    }
})();