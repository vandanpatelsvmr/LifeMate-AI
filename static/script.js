/**
 * LifeMate AI Client Application Controller
 * Handles SPA navigation, AJAX requests, form submissions, and UI renders
 */

document.addEventListener("DOMContentLoaded", () => {
    // Check API Key status on launch
    checkApiKeyStatus();

    // ── SPA Routing Inside Dashboard ──────────────────────
    const menuLinks = document.querySelectorAll(".menu-link");
    const viewSections = document.querySelectorAll(".view-section");

    menuLinks.forEach(link => {
        link.addEventListener("click", (e) => {
            e.preventDefault();
            const targetSectionId = link.getAttribute("data-target");
            
            // Remove active classes
            menuLinks.forEach(l => l.classList.remove("active"));
            viewSections.forEach(s => s.classList.remove("active"));
            
            // Add active classes
            link.classList.add("active");
            const targetSection = document.getElementById(targetSectionId);
            if (targetSection) {
                targetSection.classList.add("active");
            }
        });
    });

    // ── API Key Management ────────────────────────────────
    const apiBadge = document.getElementById("api-badge");
    const keyModal = document.getElementById("key-modal");
    const closeModal = document.getElementById("close-modal");
    const cancelKeyBtn = document.getElementById("cancel-key-btn");
    const keyForm = document.getElementById("api-key-form");
    const clearKeyBtn = document.getElementById("clear-key-btn");

    const showModal = () => keyModal.classList.add("active");
    const hideModal = () => keyModal.classList.remove("active");

    if (apiBadge) apiBadge.addEventListener("click", showModal);
    if (closeModal) closeModal.addEventListener("click", hideModal);
    if (cancelKeyBtn) cancelKeyBtn.addEventListener("click", hideModal);

    // Save API key
    if (keyForm) {
        keyForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const apiKeyInput = document.getElementById("api-key-input").value.trim();
            if (!apiKeyInput) return;

            try {
                const response = await fetch("/api/set_key", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ api_key: apiKeyInput })
                });
                const result = await response.json();
                if (result.success) {
                    hideModal();
                    updateApiIndicator(true);
                    alert("Gemini API Key saved for this session!");
                } else {
                    alert("Error: " + result.error);
                }
            } catch (err) {
                console.error("Error setting key:", err);
                alert("Failed to save API key. Check console.");
            }
        });
    }

    // Clear API Key
    if (clearKeyBtn) {
        clearKeyBtn.addEventListener("click", async () => {
            try {
                const response = await fetch("/api/clear_key", { method: "POST" });
                const result = await response.json();
                if (result.success) {
                    document.getElementById("api-key-input").value = "";
                    hideModal();
                    updateApiIndicator(false);
                    alert("API Key cleared.");
                }
            } catch (err) {
                console.error("Error clearing key:", err);
            }
        });
    }

    // ── Forms Submission & Generation handlers ────────────
    setupFormHandler("planner-form", "planner", {
        tasks: () => document.getElementById("planner-tasks").value,
        available_time: () => document.getElementById("planner-time").value,
        wellness_goals: () => document.getElementById("planner-wellness").value
    });

    setupFormHandler("budget-form", "budget", {
        income: () => document.getElementById("budget-income").value,
        expenses: () => document.getElementById("budget-expenses").value,
        savings_goal: () => document.getElementById("budget-savings").value
    });

    setupFormHandler("wellness-form", "wellness", {
        age: () => document.getElementById("wellness-age").value,
        lifestyle: () => document.getElementById("wellness-lifestyle").value,
        goals: () => document.getElementById("wellness-goals").value,
        diet_pref: () => document.getElementById("wellness-diet").value
    });

    setupFormHandler("community-form", "community", {
        question: () => document.getElementById("community-question").value
    });

    setupFormHandler("emergency-form", "emergency", {
        emergency_type: () => document.getElementById("emergency-type").value,
        location: () => document.getElementById("emergency-location").value
    });

    // ── AI Smart Assistant Chat Mode ──────────────────────
    const chatForm = document.getElementById("chat-form");
    const chatInput = document.getElementById("chat-input");
    const chatMessages = document.getElementById("chat-messages");

    if (chatForm) {
        chatForm.addEventListener("submit", async (e) => {
            e.preventDefault();
            const messageText = chatInput.value.trim();
            if (!messageText) return;

            // Add User message bubble
            appendMessage(messageText, "user");
            chatInput.value = "";

            // Add typing placeholder
            const typingBubble = appendMessage("<div class='spinner' style='width:20px;height:20px;margin-bottom:0;'></div> Thinking...", "ai");

            try {
                const response = await fetch("/api/generate", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        feature: "chat",
                        data: { message: messageText }
                    })
                });
                const data = await response.json();
                typingBubble.remove();

                if (data.success) {
                    appendMessage(data.result, "ai", true);
                } else {
                    if (data.needs_key) {
                        appendMessage("⚠️ Please configure your Gemini API key in the top settings panel to use the Chat Assistant.", "ai");
                        showModal();
                    } else {
                        appendMessage("⚠️ Error: " + data.error, "ai");
                    }
                }
            } catch (err) {
                typingBubble.remove();
                appendMessage("⚠️ Network error. Please try again.", "ai");
                console.error(err);
            }
        });
    }

    function appendMessage(content, sender, isMarkdown = false) {
        const bubble = document.createElement("div");
        bubble.className = `message-bubble ${sender}`;
        
        if (isMarkdown) {
            bubble.classList.add("markdown-body");
            bubble.innerHTML = parseMarkdown(content);
        } else {
            bubble.innerHTML = content;
        }
        
        chatMessages.appendChild(bubble);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        return bubble;
    }

    // ── Action Buttons (Copy / Export) ──────────────────
    document.querySelectorAll(".view-section").forEach(section => {
        const copyBtn = section.querySelector(".btn-copy");
        const printBtn = section.querySelector(".btn-print");
        const outputDiv = section.querySelector(".output-content");
        
        if (copyBtn) {
            copyBtn.addEventListener("click", () => {
                if (outputDiv && outputDiv.innerText) {
                    navigator.clipboard.writeText(outputDiv.innerText)
                        .then(() => alert("Copied to clipboard!"))
                        .catch(() => alert("Failed to copy."));
                }
            });
        }

        if (printBtn) {
            printBtn.addEventListener("click", () => {
                if (outputDiv && outputDiv.getAttribute("data-raw-content")) {
                    const rawContent = encodeURIComponent(outputDiv.getAttribute("data-raw-content"));
                    const title = encodeURIComponent(section.querySelector(".output-title").innerText || "LifeMate AI Guide");
                    window.open(`/result?content=${rawContent}&title=${title}`, "_blank");
                }
            });
        }
    });
});

// Helper: Setup submission for modular forms
function setupFormHandler(formId, featureName, dataExtractors) {
    const form = document.getElementById(formId);
    if (!form) return;

    const section = form.closest(".view-section");
    const placeholder = section.querySelector(".output-placeholder");
    const loader = section.querySelector(".loading-box");
    const outputDiv = section.querySelector(".output-content");

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        // 1. Extract inputs
        const requestData = {};
        for (const [key, extractor] of Object.entries(dataExtractors)) {
            requestData[key] = extractor();
        }

        // 2. Adjust visibility
        placeholder.style.display = "none";
        outputDiv.style.display = "none";
        loader.style.display = "flex";

        try {
            const response = await fetch("/api/generate", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    feature: featureName,
                    data: requestData
                })
            });
            const data = await response.json();
            
            loader.style.display = "none";

            if (data.success) {
                outputDiv.style.display = "block";
                outputDiv.setAttribute("data-raw-content", data.result);
                outputDiv.innerHTML = parseMarkdown(data.result);
            } else {
                placeholder.style.display = "flex";
                if (data.needs_key) {
                    alert("Please enter your Gemini API Key in the top badge panel to run this assistant.");
                    document.getElementById("key-modal").classList.add("active");
                } else {
                    alert("Error: " + data.error);
                }
            }
        } catch (err) {
            loader.style.display = "none";
            placeholder.style.display = "flex";
            alert("API connection failed. Make sure your Python Flask backend is running.");
            console.error(err);
        }
    });
}

// ── Check API Key Status ──────────────────────────────────
async function checkApiKeyStatus() {
    try {
        const response = await fetch("/api/check_key");
        const data = await response.json();
        updateApiIndicator(data.configured);
    } catch (err) {
        console.error("Failed to query API key status:", err);
        updateApiIndicator(false);
    }
}

function updateApiIndicator(isConfigured) {
    const textSpan = document.getElementById("api-status-text");
    const dot = document.getElementById("api-status-dot");
    
    if (isConfigured) {
        textSpan.innerText = "Gemini AI: Connected";
        dot.className = "indicator-dot active";
    } else {
        textSpan.innerText = "Gemini AI: Setup Key";
        dot.className = "indicator-dot inactive";
    }
}

// ── Simple Client-Side Markdown Parser ──────────────────────
function parseMarkdown(md) {
    if (!md) return "";
    
    // Safety escape HTML tags to prevent XSS
    let html = md
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    // Headings
    html = html.replace(/^### (.*?)$/gm, "<h3>$1</h3>");
    html = html.replace(/^## (.*?)$/gm, "<h2>$1</h2>");
    html = html.replace(/^# (.*?)$/gm, "<h1>$1</h1>");

    // Blockquotes
    html = html.replace(/^&gt; (.*?)$/gm, "<blockquote>$1</blockquote>");

    // Bold text
    html = html.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");

    // Bullet Lists (supporting nested lists is complex, this handles simple flat ones)
    // Convert - or * lists
    let inList = false;
    const lines = html.split("\n");
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        if (line.startsWith("- ") || line.startsWith("* ")) {
            let itemContent = line.substring(2);
            if (!inList) {
                lines[i] = "<ul><li>" + itemContent + "</li>";
                inList = true;
            } else {
                lines[i] = "<li>" + itemContent + "</li>";
            }
        } else {
            if (inList) {
                lines[i-1] = lines[i-1] + "</ul>";
                inList = false;
            }
        }
    }
    if (inList) {
        lines[lines.length - 1] = lines[lines.length - 1] + "</ul>";
    }
    html = lines.join("\n");

    // Convert newlines (paragraphs)
    html = html.replace(/\n\n/g, "</p><p>");
    
    // Wrap text inside body with p tags if not in headers or list tags
    // Simple way: wrap whole text and clean up double tags
    html = "<p>" + html + "</p>";
    html = html
        .replace(/<p><h3>/g, "<h3>")
        .replace(/<\/h3><\/p>/g, "</h3>")
        .replace(/<p><h2>/g, "<h2>")
        .replace(/<\/h2><\/p>/g, "</h2>")
        .replace(/<p><h1>/g, "<h1>")
        .replace(/<\/h1><\/p>/g, "</h1>")
        .replace(/<p><ul>/g, "<ul>")
        .replace(/<\/ul><\/p>/g, "</ul>")
        .replace(/<p><blockquote>/g, "<blockquote>")
        .replace(/<\/blockquote><\/p>/g, "</blockquote>")
        .replace(/<p><\/p>/g, ""); // Clean up empties

    return html;
}
