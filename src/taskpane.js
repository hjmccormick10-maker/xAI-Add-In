/* taskpane.js */
import './taskpane.css'; // Ensure CSS import for task pane styling

// Log script load
console.log('DEBUG: taskpane.js loaded at ' + new Date().toISOString());

// Debounce function
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        return new Promise((resolve, reject) => {
            const later = () => {
                clearTimeout(timeout);
                try {
                    const result = func(...args);
                    resolve(result);
                } catch (error) {
                    reject(error);
                }
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        });
    };
}

// Global state
let isApiKeyValid = false;
let isSubmitting = false;
let isEventHandlersInitialized = false;
let currentModel = null;
let isFirstSubmission = true;
let currentSystemPrompt = 'You are a concise assistant for proposal writing in a Word add-in.';
let currentPersona = '';
let currentStyle = 'Normal';

// Prevent multiple initializations
if (window.__xaiAddInInitialized) {
    console.log('DEBUG: Skipped taskpane.js initialization - already initialized');
} else {
    window.__xaiAddInInitialized = true;

    function cleanText(text) {
        if (!text || typeof text !== 'string') return '';
        return text
            .replace(/\\n\\r/g, '\r\n')
            .replace(/\\n/g, '\r\n')
            .replace(/\\r/g, '\r\n')
            .replace(/\\\\n/g, '\r\n')
            .replace(/\n/g, '\r\n')
            .replace(/\\t/g, '\t')
            .replace(/\\\\/g, '\\')
            .replace(/\\&/g, '&')
            .replace(/\\\*/g, '*')
            .replace(/\*\*(.*?)\*\*/g, '$1')
            .replace(/\*(.*?)\*/g, '$1')
            .replace(/^#+\s*(.*?)$/gm, '$1')
            .replace(/^\s*[-*]\s*(.*?)$/gm, '$1')
            .replace(/\\[^nr&t*]/g, '')
            .trim();
    }

    function isInitializedCheck() {
        if (!window.__xaiAddInInitialized || !Office.context) {
            console.log('isInitializedCheck: Office not initialized yet');
            const statusMessage = document.getElementById('statusMessage');
            if (statusMessage) statusMessage.textContent = 'Office is not initialized. Please wait and try again.';
            return false;
        }
        return true;
    }

    function ensureButtonVisibility() {
        const buttons = [
            document.getElementById('enterApiKeyButton'),
            document.getElementById('enterModelButton'),
            document.getElementById('configureChatButton'),
        ];
        buttons.forEach(button => {
            if (button && button.style.display !== 'block') {
                button.style.display = 'block';
                console.log(`ensureButtonVisibility: Set ${button.id} to display: block`);
            }
        });
        setTimeout(() => {
            buttons.forEach(button => {
                if (button && button.style.display !== 'block') {
                    console.warn(`ensureButtonVisibility: ${button.id} was hidden, forcing display: block`);
                    button.style.display = 'block';
                }
            });
            const responseBox = document.getElementById('responseBox');
            if (responseBox && responseBox.style.display !== 'block') {
                console.warn('ensureButtonVisibility: responseBox was hidden, forcing display: block');
                responseBox.style.display = 'block';
            }
        }, 1000);
    }

    async function checkConfigOnLoad() {
        const timerId = `checkConfigOnLoad_${Date.now()}`;
        console.time(timerId);
        const statusMessage = document.getElementById('statusMessage');
        const apiKeyModal = document.getElementById('apiKeyModal');
        const modelModal = document.getElementById('modelModal');
        const styleInput = document.getElementById('styleInput');
        if (!statusMessage || !apiKeyModal || !modelModal || !styleInput) {
            console.error('checkConfigOnLoad: Missing DOM elements');
            if (statusMessage) statusMessage.textContent = 'Error: UI elements missing.';
            console.timeEnd(timerId);
            return;
        }
        try {
            const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            const configUrl = isLocal
                ? 'http://localhost:3001/get-config'
                : 'xai_config.json'; // Updated to lowercase
            const response = await fetch(configUrl, {
                method: 'GET',
                headers: { 'Content-Type': 'application/json' },
                mode: 'cors',
            });
            if (!response.ok) throw new Error(`Failed to fetch config, status: ${response.status}`);
            const config = await response.json();
            console.log('checkConfigOnLoad: Config received', config);

            // Load API key from roamingSettings in production
            let apiKey = null;
            if (!isLocal) {
                apiKey = Office.context.roamingSettings.get('xaiApiKey');
                isApiKeyValid = !!apiKey && apiKey.startsWith('xai-') && apiKey.length >= 20;
            } else {
                isApiKeyValid = config.hasValidApiKey || false;
            }

            // Handle token limit
            document.getElementsByName('tokenLimit').forEach(input => {
                input.checked = parseInt(input.value) === (config.XAI_MAX_TOKENS || 1000);
            });

            // Handle model
            if (config.XAI_MODEL) {
                // Assume model is valid; adjust if xAI provides a validation endpoint
                currentModel = config.XAI_MODEL;
                localStorage.setItem('xaiModelVersion', currentModel);
                document.getElementById('modelName').textContent = currentModel;
                document.getElementById('modelInput').value = currentModel;
            } else {
                console.log('checkConfigOnLoad: No model in config');
                currentModel = null;
                document.getElementById('modelName').textContent = 'unknown';
                statusMessage.textContent = 'No model configured. Please select a model.';
                modelModal.style.display = 'block';
                document.getElementById('modelInput')?.focus();
            }

            // Handle system prompt
            currentSystemPrompt = config.XAI_SYSTEM_PROMPT || currentSystemPrompt;
            localStorage.setItem('xaiSystemPrompt', currentSystemPrompt);
            document.getElementById('systemPromptInput').value = currentSystemPrompt;
            document.getElementById('userPersona').value = currentPersona;

            // Handle style (XAI_FONT)
            if (config.XAI_FONT) {
                const cleanedStyle = cleanText(config.XAI_FONT).trim();
                console.log(`checkConfigOnLoad: Validating style "${cleanedStyle}" from XAI_FONT`);
                try {
                    await Word.run(async (context) => {
                        const styles = context.document.getStyles();
                        context.load(styles);
                        await context.sync();
                        const styleExists = styles.items.some(style => style.nameLocal.toLowerCase() === cleanedStyle.toLowerCase());
                        if (styleExists) {
                            currentStyle = styles.items.find(style => style.nameLocal.toLowerCase() === cleanedStyle.toLowerCase()).nameLocal;
                            styleInput.value = currentStyle;
                            localStorage.setItem('xaiStyle', currentStyle);
                            if (!isLocal) Office.context.roamingSettings.set('xaiStyle', currentStyle);
                            console.log(`checkConfigOnLoad: Set style to "${currentStyle}" from XAI_FONT`);
                        } else {
                            console.warn(`checkConfigOnLoad: Invalid style "${cleanedStyle}" in XAI_FONT`);
                            currentStyle = 'Normal';
                            styleInput.value = currentStyle;
                            localStorage.setItem('xaiStyle', currentStyle);
                            if (!isLocal) Office.context.roamingSettings.set('xaiStyle', currentStyle);
                            statusMessage.textContent = "Invalid style in configuration. Defaulted to 'Normal'.";
                        }
                        if (!isLocal) await Office.context.roamingSettings.saveAsync();
                    });
                } catch (error) {
                    console.error(`checkConfigOnLoad: Error validating style "${cleanedStyle}"`, error.message);
                    currentStyle = 'Normal';
                    styleInput.value = currentStyle;
                    localStorage.setItem('xaiStyle', currentStyle);
                    if (!isLocal) {
                        Office.context.roamingSettings.set('xaiStyle', currentStyle);
                        await Office.context.roamingSettings.saveAsync();
                    }
                    statusMessage.textContent = `Error validating style: ${error.message}. Defaulted to 'Normal'.`;
                }
            } else {
                console.log('checkConfigOnLoad: No style in config (XAI_FONT missing)');
                currentStyle = 'Normal';
                styleInput.value = currentStyle;
                localStorage.setItem('xaiStyle', currentStyle);
                if (!isLocal) {
                    Office.context.roamingSettings.set('xaiStyle', currentStyle);
                    await Office.context.roamingSettings.saveAsync();
                }
                statusMessage.textContent = "No style configured. Defaulted to 'Normal'.";
            }

            // Handle API key
            if (!isApiKeyValid) {
                console.log('checkConfigOnLoad: No valid API key');
                localStorage.removeItem('xaiApiKey');
                statusMessage.textContent = 'No valid API key found. Please enter a key.';
                apiKeyModal.style.display = 'block';
                apiKeyModal.focus();
            } else {
                console.log('checkConfigOnLoad: Valid API key found');
                localStorage.setItem('xaiApiKey', '[REDACTED]');
                statusMessage.textContent = 'API key loaded successfully.';
            }

            // Final status
            if (isApiKeyValid && currentModel && currentStyle) {
                statusMessage.textContent = 'Ready to submit questions.';
            }
        } catch (error) {
            console.error('checkConfigOnLoad: Error', error.message);
            localStorage.removeItem('xaiApiKey');
            localStorage.removeItem('xaiModelVersion');
            localStorage.removeItem('xaiSystemPrompt');
            localStorage.removeItem('xaiPersona');
            localStorage.removeItem('xaiStyle');
            isApiKeyValid = false;
            currentModel = null;
            currentStyle = 'Normal';
            styleInput.value = currentStyle;
            document.getElementById('modelName').textContent = 'unknown';
            statusMessage.textContent = `Error loading configuration: ${error.message}. Please enter an API key, model, and style.`;
            apiKeyModal.style.display = 'block';
            apiKeyModal.focus();
        } finally {
            console.timeEnd(timerId);
            ensureButtonVisibility();
        }
    }

    const saveModel = debounce(async function () {
        const modelInput = document.getElementById('modelInput');
        const statusMessage = document.getElementById('statusMessage');
        const modelModal = document.getElementById('modelModal');
        if (!modelInput || !statusMessage || !modelModal) {
            console.error('saveModel: Missing DOM elements');
            statusMessage.textContent = 'Error: Model input UI missing.';
            return;
        }
        const newModel = cleanText(modelInput.value);
        if (!newModel) {
            console.log('saveModel: No model entered');
            statusMessage.textContent = 'Please enter a model name.';
            return;
        }
        try {
            const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            if (isLocal) {
                const saveResponse = await fetch('http://localhost:3001/save-env', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ model: newModel }),
                    mode: 'cors',
                });
                if (!saveResponse.ok) throw new Error(`Failed to save model, status: ${saveResponse.status}`);
            } else {
                Office.context.roamingSettings.set('xaiModelVersion', newModel);
                await Office.context.roamingSettings.saveAsync();
            }
            currentModel = newModel;
            localStorage.setItem('xaiModelVersion', newModel);
            document.getElementById('modelName').textContent = newModel;
            document.getElementById('modelInput').value = newModel;
            statusMessage.textContent = 'Model updated successfully.';
            modelModal.style.display = 'none';
            console.log(`saveModel: Set model to "${newModel}"`);
        } catch (error) {
            console.error('saveModel: Error', error.message);
            statusMessage.textContent = `Error saving model: ${error.message}`;
        } finally {
            ensureButtonVisibility();
        }
    }, 500);

    const configureChat = debounce(async function () {
        console.log('configureChat: Starting');
        const systemPromptInput = document.getElementById('systemPromptInput');
        const userPersona = document.getElementById('userPersona');
        const userMessageInput = document.getElementById('userMessageInput');
        const statusMessage = document.getElementById('statusMessage');
        const chatConfigModal = document.getElementById('chatConfigModal');
        if (!systemPromptInput || !userPersona || !userMessageInput || !statusMessage || !chatConfigModal) {
            console.error('configureChat: Missing DOM elements');
            statusMessage.textContent = 'Error: Chat configuration UI missing.';
            return;
        }
        systemPromptInput.value = currentSystemPrompt || '';
        userPersona.value = currentPersona || '';
        userMessageInput.value = '';
        chatConfigModal.style.display = 'block';
        systemPromptInput.focus();
    }, 500);

    const saveChatConfig = debounce(async function () {
        console.log('saveChatConfig: Starting');
        const systemPromptInput = document.getElementById('systemPromptInput');
        const userPersona = document.getElementById('userPersona');
        const userMessageInput = document.getElementById('userMessageInput');
        const statusMessage = document.getElementById('statusMessage');
        const chatConfigModal = document.getElementById('chatConfigModal');
        if (!systemPromptInput || !userPersona || !userMessageInput || !statusMessage || !chatConfigModal) {
            console.error('saveChatConfig: Missing DOM elements');
            statusMessage.textContent = 'Error: Chat configuration UI missing.';
            return;
        }
        const newSystemPrompt = cleanText(systemPromptInput.value);
        const newPersona = cleanText(userPersona.value);
        const newUserMessage = cleanText(userMessageInput.value);
        if (!newSystemPrompt) {
            console.log('saveChatConfig: No system prompt entered');
            statusMessage.textContent = 'Please enter a system prompt.';
            systemPromptInput.focus();
            return;
        }
        if (newPersona.length > 50) {
            console.log('saveChatConfig: Persona too long');
            statusMessage.textContent = 'Persona must be 50 characters or less.';
            userPersona.focus();
            return;
        }
        try {
            const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            if (isLocal) {
                const saveResponse = await fetch('http://localhost:3001/save-env', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ systemPrompt: newSystemPrompt }),
                    mode: 'cors',
                });
                if (!saveResponse.ok) throw new Error(`Failed to save chat config, status: ${saveResponse.status}`);
            } else {
                Office.context.roamingSettings.set('xaiSystemPrompt', newSystemPrompt);
                Office.context.roamingSettings.set('xaiPersona', newPersona);
                await Office.context.roamingSettings.saveAsync();
            }
            currentSystemPrompt = newSystemPrompt;
            localStorage.setItem('xaiSystemPrompt', newSystemPrompt);
            currentPersona = newPersona;
            localStorage.setItem('xaiPersona', newPersona);
            const questionBox = document.getElementById('questionBox');
            if (questionBox) {
                questionBox.value = newPersona && newUserMessage ? `As a ${newPersona}: ${newUserMessage}` : newUserMessage || '';
            }
            statusMessage.textContent = 'Chat configuration saved successfully.';
            chatConfigModal.style.display = 'none';
            console.log(`saveChatConfig: Set system prompt to "${newSystemPrompt}", persona to "${newPersona}"`);
        } catch (error) {
            console.error('saveChatConfig: Error', error.message);
            statusMessage.textContent = `Error saving chat configuration: ${error.message}`;
            systemPromptInput.value = currentSystemPrompt;
            userPersona.value = currentPersona;
            userMessageInput.value = '';
        } finally {
            ensureButtonVisibility();
        }
    }, 500);

    function cancelChatConfig() {
        console.log('cancelChatConfig: Starting');
        const chatConfigModal = document.getElementById('chatConfigModal');
        const systemPromptInput = document.getElementById('systemPromptInput');
        const userPersona = document.getElementById('userPersona');
        const userMessageInput = document.getElementById('userMessageInput');
        const statusMessage = document.getElementById('statusMessage');
        if (!chatConfigModal) {
            console.error('cancelChatConfig: chatConfigModal not found in DOM');
            statusMessage.textContent = 'Error: Chat configuration UI missing.';
            return;
        }
        chatConfigModal.style.display = 'none';
        if (systemPromptInput) systemPromptInput.value = currentSystemPrompt;
        if (userPersona) userPersona.value = currentPersona;
        if (userMessageInput) userMessageInput.value = '';
        statusMessage.textContent = 'Chat configuration cancelled.';
        ensureButtonVisibility();
    }

    const saveApiKey = debounce(async function () {
        const apiKeyInput = document.getElementById('apiKeyInput');
        const statusMessage = document.getElementById('statusMessage');
        const apiKeyModal = document.getElementById('apiKeyModal');
        if (!apiKeyInput || !statusMessage || !apiKeyModal) {
            console.error('saveApiKey: Missing DOM elements');
            statusMessage.textContent = 'Error: API key input UI missing.';
            return;
        }
        const apiKey = apiKeyInput.value.trim();
        if (!apiKey || !apiKey.startsWith('xai-') || apiKey.length < 20 || !/^[a-zA-Z0-9-]+$/.test(apiKey)) {
            console.log('saveApiKey: Invalid API key format');
            isApiKeyValid = false;
            statusMessage.textContent = "Invalid API key format. Must start with 'xai-' and be at least 20 characters.";
            apiKeyInput.focus();
            return;
        }
        try {
            const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            if (!isLocal) {
                // Test API key with a minimal request
                const testResponse = await fetch('https://api.x.ai/v1/chat/completions', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        Authorization: `Bearer ${apiKey}`,
                    },
                    body: JSON.stringify({
                        messages: [{ role: 'user', content: 'Test' }],
                        model: 'grok-3',
                        max_tokens: 10,
                    }),
                    mode: 'cors',
                });
                if (!testResponse.ok) {
                    const errorData = await testResponse.json();
                    throw new Error(errorData.error || 'API key validation failed');
                }
                Office.context.roamingSettings.set('xaiApiKey', apiKey);
                await Office.context.roamingSettings.saveAsync();
            } else {
                const saveResponse = await fetch('http://localhost:3001/save-env', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ apiKey }),
                    mode: 'cors',
                });
                if (!saveResponse.ok) {
                    const errorData = await saveResponse.json();
                    throw new Error(errorData.error || 'Failed to save API key to server');
                }
            }
            localStorage.setItem('xaiApiKey', '[REDACTED]');
            isApiKeyValid = true;
            statusMessage.textContent = 'API key saved and validated successfully.';
            apiKeyModal.style.display = 'none';
            console.log('saveApiKey: API key saved');
        } catch (error) {
            console.error('saveApiKey: Error', error.message);
            isApiKeyValid = false;
            statusMessage.textContent = `Error saving API key: ${error.message}`;
            apiKeyInput.focus();
        } finally {
            ensureButtonVisibility();
        }
    }, 500);

    function cancelQuestion() {
        const questionBox = document.getElementById('questionBox');
        const statusMessage = document.getElementById('statusMessage');
        const tokenDisplay = document.getElementById('tokenDisplay');
        if (questionBox) questionBox.value = '';
        if (statusMessage) statusMessage.textContent = 'Question and response cleared.';
        if (tokenDisplay) tokenDisplay.style.display = 'none';
        console.log('cancelQuestion: Question discarded');
    }

    function discardResponse() {
        const responseBox = document.getElementById('responseBox');
        const tokenDisplay = document.getElementById('tokenDisplay');
        const statusMessage = document.getElementById('statusMessage');
        if (responseBox) {
            responseBox.value = '';
            responseBox.style.height = 'auto';
        }
        if (tokenDisplay) tokenDisplay.style.display = 'none';
        if (statusMessage) statusMessage.textContent = 'Response discarded.';
        console.log('discardResponse: Response discarded');
    }

    function toggleApiKeyVisibility() {
        const apiKeyInput = document.getElementById('apiKeyInput');
        const showApiKeyCheckbox = document.getElementById('showApiKeyCheckbox');
        if (apiKeyInput && showApiKeyCheckbox) {
            apiKeyInput.type = showApiKeyCheckbox.checked ? 'text' : 'password';
        }
    }

    async function insertResponse() {
        console.log('insertResponse: Starting');
        if (!isInitializedCheck()) return;
        const responseBox = document.getElementById('responseBox');
        const statusMessage = document.getElementById('statusMessage');
        if (!responseBox) {
            console.error('insertResponse: Missing DOM elements', { responseBox });
            if (statusMessage) statusMessage.textContent = 'Error: UI elements missing.';
            return;
        }
        if (!responseBox.value) {
            console.log('insertResponse: No response to insert');
            if (statusMessage) statusMessage.textContent = 'No response to insert.';
            return;
        }
        try {
            await Word.run(async (context) => {
                const wordDocument = context.document;
                const styles = wordDocument.getStyles();
                const selection = wordDocument.getSelection();
                context.load(styles);
                context.load(selection);
                await context.sync();
                console.log('insertResponse: Available styles =', styles.items.map(s => s.nameLocal));
                const styleName = currentStyle;
                console.log('insertResponse: Using style =', styleName);
                const cleanedText = cleanText(responseBox.value);
                const paragraphs = cleanedText.split('\r\n').filter(p => p.trim());
                for (let i = paragraphs.length - 1; i >= 0; i--) {
                    const paraText = paragraphs[i];
                    console.log(`insertResponse: Preparing paragraph ${i + 1}:`, paraText);
                    const paragraph = selection.insertParagraph(paraText, Word.InsertLocation.after);
                    try {
                        paragraph.style = styleName;
                    } catch (error) {
                        console.log(`insertResponse: Failed to apply style ${styleName}: ${error.message}`);
                        paragraph.style = 'Normal';
                        paragraph.font.name = 'Arial';
                        paragraph.font.size = 11;
                        paragraph.font.color = '#1A3C5A';
                    }
                    if (styleName === 'Normal') {
                        paragraph.font.name = 'Arial';
                        paragraph.font.size = 11;
                        paragraph.font.color = '#1A3C5A';
                    }
                    if (paraText.match(/\*\*(.*?)\*\*/)) paragraph.font.bold = true;
                    if (paraText.match(/^\s*[-*]\s*(.*?)$/)) paragraph.leftIndent = 20;
                    if (paraText.match(/^[A-Za-z\s]+:$/) || paraText.match(/^#+\s*(.*?)$/)) {
                        paragraph.font.size = 14;
                        paragraph.font.bold = true;
                        paragraph.spaceAfter = 6;
                    }
                    paragraph.spaceAfter = 6;
                }
                await context.sync();
                console.log('insertResponse: All paragraphs inserted');
                if (statusMessage) statusMessage.textContent = 'Response inserted.';
            });
        } catch (error) {
            console.error('insertResponse: Error', error.message, error.stack);
            if (statusMessage) statusMessage.textContent = `Error inserting response: ${error.message}`;
        } finally {
            ensureButtonVisibility();
        }
    }

    async function insertSnippet() {
        console.log('insertSnippet: Starting');
        if (!isInitializedCheck()) return;
        const responseBox = document.getElementById('responseBox');
        const statusMessage = document.getElementById('statusMessage');
        if (!responseBox) {
            console.error('insertSnippet: Missing DOM elements', { responseBox });
            if (statusMessage) statusMessage.textContent = 'Error: UI elements missing.';
            return;
        }
        const selectionStart = responseBox.selectionStart;
        const selectionEnd = responseBox.selectionEnd;
        const selectedText = cleanText(responseBox.value.substring(selectionStart, selectionEnd) || '');
        if (!selectedText) {
            console.log('insertSnippet: No text selected');
            if (statusMessage) statusMessage.textContent = 'Please select text to insert.';
            return;
        }
        try {
            await Word.run(async (context) => {
                const wordDocument = context.document;
                const styles = wordDocument.getStyles();
                const selection = wordDocument.getSelection();
                context.load(styles);
                context.load(selection);
                await context.sync();
                console.log('insertSnippet: Available styles =', styles.items.map(s => s.nameLocal));
                const styleName = currentStyle;
                const paragraphs = selectedText.split('\r\n').filter(p => p.trim());
                for (let i = paragraphs.length - 1; i >= 0; i--) {
                    const paraText = paragraphs[i];
                    console.log(`insertSnippet: Preparing paragraph ${i + 1}:`, paraText);
                    const paragraph = selection.insertParagraph(paraText, Word.InsertLocation.after);
                    try {
                        paragraph.style = styleName;
                    } catch (error) {
                        console.log(`insertSnippet: Failed to apply style ${styleName}: ${error.message}`);
                        paragraph.style = 'Normal';
                        paragraph.font.name = 'Arial';
                        paragraph.font.size = 11;
                        paragraph.font.color = '#1A3C5A';
                    }
                    if (styleName === 'Normal') {
                        paragraph.font.name = 'Arial';
                        paragraph.font.size = 11;
                        paragraph.font.color = '#1A3C5A';
                    }
                    if (paraText.match(/\*\*(.*?)\*\*/)) paragraph.font.bold = true;
                    if (paraText.match(/^\s*[-*]\s*(.*?)$/)) paragraph.leftIndent = 20;
                    if (paraText.match(/^[A-Za-z\s]+:$/) || paraText.match(/^#+\s*(.*?)$/)) {
                        paragraph.font.size = 14;
                        paragraph.font.bold = true;
                        paragraph.spaceAfter = 6;
                    }
                    paragraph.spaceAfter = 6;
                }
                await context.sync();
                console.log('insertSnippet: All paragraphs inserted');
                if (statusMessage) statusMessage.textContent = 'Snippet inserted.';
            });
            responseBox.setSelectionRange(selectionStart, selectionEnd);
            responseBox.focus();
        } catch (error) {
            console.error('insertSnippet: Error', error.message, error.stack);
            if (statusMessage) statusMessage.textContent = `Error inserting snippet: ${error.message}`;
        } finally {
            ensureButtonVisibility();
        }
    }

    const saveStyle = debounce(async function () {
        const styleModalInput = document.getElementById('styleModalInput');
        const styleInput = document.getElementById('styleInput');
        const statusMessage = document.getElementById('statusMessage');
        const styleModal = document.getElementById('styleModal');
        if (!styleModalInput || !styleInput || !statusMessage || !styleModal) {
            console.error('saveStyle: Missing DOM elements');
            statusMessage.textContent = 'Error: Style input UI missing.';
            return;
        }
        const newStyle = cleanText(styleModalInput.value).trim();
        if (!newStyle) {
            console.log('saveStyle: No style entered');
            statusMessage.textContent = 'Please enter a style name.';
            styleModalInput.focus();
            return;
        }
        try {
            await Word.run(async (context) => {
                const styles = context.document.getStyles();
                context.load(styles);
                await context.sync();
                const styleExists = styles.items.some(style => style.nameLocal.toLowerCase() === newStyle.toLowerCase());
                if (!styleExists) {
                    console.warn('saveStyle: Style not found', newStyle);
                    statusMessage.textContent = 'Style is not available.';
                    currentStyle = 'Normal';
                    styleInput.value = currentStyle;
                    localStorage.setItem('xaiStyle', currentStyle);
                } else {
                    currentStyle = styles.items.find(style => style.nameLocal.toLowerCase() === newStyle.toLowerCase()).nameLocal;
                    styleInput.value = currentStyle;
                    localStorage.setItem('xaiStyle', currentStyle);
                    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
                    if (!isLocal) {
                        Office.context.roamingSettings.set('xaiStyle', currentStyle);
                        await Office.context.roamingSettings.saveAsync();
                    } else {
                        const saveResponse = await fetch('http://localhost:3001/save-env', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ font: currentStyle }),
                            mode: 'cors',
                        });
                        if (!saveResponse.ok) throw new Error(`Failed to save style, status: ${saveResponse.status}`);
                    }
                    statusMessage.textContent = `Style set to "${currentStyle}".`;
                }
                styleModal.style.display = 'none';
                console.log(`saveStyle: Set style to "${currentStyle}"`);
            });
        } catch (error) {
            console.error('saveStyle: Error', error.message);
            statusMessage.textContent = `Error saving style: ${error.message}. Defaulted to 'Normal'.`;
            currentStyle = 'Normal';
            styleInput.value = currentStyle;
            localStorage.setItem('xaiStyle', currentStyle);
            if (!isLocal) {
                Office.context.roamingSettings.set('xaiStyle', currentStyle);
                await Office.context.roamingSettings.saveAsync();
            }
            styleModal.style.display = 'none';
        }
    }, 500);

    function cancelStyle() {
        const styleModal = document.getElementById('styleModal');
        const styleModalInput = document.getElementById('styleModalInput');
        const statusMessage = document.getElementById('statusMessage');
        if (!styleModal || !styleModalInput) {
            console.error('cancelStyle: Missing DOM elements');
            statusMessage.textContent = 'Error: Style input UI missing.';
            return;
        }
        styleModalInput.value = currentStyle;
        styleModal.style.display = 'none';
        statusMessage.textContent = 'Style change cancelled.';
        console.log('cancelStyle: Style change cancelled');
    }

    function handleQuestionBoxKeydown(event) {
        if (event.key === 'Enter' && !event.shiftKey && !isSubmitting) {
            event.preventDefault();
            console.log('handleQuestionBoxKeydown: Enter key pressed, triggering submitQuestion');
            submitQuestion();
        }
    }

    const submitQuestion = debounce(async function () {
        if (isSubmitting) {
            console.log('submitQuestion: Submission debounced, already in progress');
            return;
        }
        isSubmitting = true;
        const timerId = `submitQuestion_${Date.now()}`;
        console.time(timerId);
        const questionBox = document.getElementById('questionBox');
        const statusMessage = document.getElementById('statusMessage');
        const apiKeyModal = document.getElementById('apiKeyModal');
        const modelModal = document.getElementById('modelModal');
        const submitButton = document.getElementById('submitButton');
        if (submitButton) {
            submitButton.disabled = true;
            submitButton.textContent = 'Submitting...';
        }
        const question = cleanText(questionBox?.value || '');
        const maxTokens = parseInt(document.querySelector('input[name="tokenLimit"]:checked')?.value || 1000);
        if (!question) {
            console.log('submitQuestion: No question entered');
            statusMessage.textContent = 'Please enter a question.';
            isSubmitting = false;
            if (submitButton) {
                submitButton.disabled = false;
                submitButton.textContent = 'Submit Question';
            }
            console.timeEnd(timerId);
            return;
        }
        if (!isApiKeyValid) {
            console.log('submitQuestion: No valid API key');
            statusMessage.textContent = 'Please enter a valid API key.';
            apiKeyModal.style.display = 'block';
            apiKeyModal.focus();
            isSubmitting = false;
            if (submitButton) {
                submitButton.disabled = false;
                submitButton.textContent = 'Submit Question';
            }
            console.timeEnd(timerId);
            return;
        }
        if (!currentModel) {
            console.log('submitQuestion: No valid model selected');
            statusMessage.textContent = 'Please select a valid model.';
            modelModal.style.display = 'block';
            document.getElementById('modelInput')?.focus();
            isSubmitting = false;
            if (submitButton) {
                submitButton.disabled = false;
                submitButton.textContent = 'Submit Question';
            }
            console.timeEnd(timerId);
            return;
        }
        try {
            const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            let apiKey = null;
            if (!isLocal) {
                apiKey = Office.context.roamingSettings.get('xaiApiKey');
            }
            statusMessage.textContent = 'Submitting question...';
            const effectiveQuestion = currentPersona && question && !question.startsWith(`As a ${currentPersona}: `)
                ? `As a ${currentPersona}: ${question}`
                : question;
            const response = await fetch(isLocal ? 'http://localhost:3001/api/chat' : 'https://api.x.ai/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    ...(isLocal ? {} : { Authorization: `Bearer ${apiKey}` }),
                },
                body: JSON.stringify({
                    messages: [
                        { role: 'system', content: currentSystemPrompt },
                        { role: 'user', content: effectiveQuestion },
                    ],
                    model: currentModel,
                    max_tokens: maxTokens,
                }),
                mode: 'cors',
            });
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `HTTP error, status: ${response.status}`);
            }
            const data = await response.json();
            const responseBox = document.getElementById('responseBox');
            if (responseBox) {
                responseBox.value = cleanText(data.choices?.[0]?.message?.content || 'No response received.');
                responseBox.style.height = 'auto';
                responseBox.style.height = `${responseBox.scrollHeight}px`;
            }
            const tokenDisplay = document.getElementById('tokenDisplay');
            if (data.usage && tokenDisplay) {
                document.getElementById('promptTokens').textContent = data.usage.prompt_tokens || 0;
                document.getElementById('completionTokens').textContent = data.usage.completion_tokens || 0;
                document.getElementById('totalTokens').textContent = data.usage.total_tokens || 0;
                tokenDisplay.style.display = 'block';
            }
            statusMessage.textContent = 'Response received.';
        } catch (error) {
            console.error('submitQuestion: Error', error.message);
            statusMessage.textContent = `Error: ${error.message}. Please check your API key or model.`;
            apiKeyModal.style.display = 'block';
            apiKeyModal.focus();
        } finally {
            isSubmitting = false;
            if (submitButton) {
                submitButton.disabled = false;
                submitButton.textContent = 'Submit Question';
            }
            console.timeEnd(timerId);
            ensureButtonVisibility();
        }
    }, 500);

    function initializeEventHandlers() {
        if (isEventHandlersInitialized) {
            console.log('DEBUG: Skipped event handler initialization - already initialized');
            return;
        }
        isEventHandlersInitialized = true;
        try {
            const elements = {
                submitButton: document.getElementById('submitButton'),
                questionBox: document.getElementById('questionBox'),
                responseBox: document.getElementById('responseBox'),
                cancelButton: document.getElementById('cancelButton'),
                enterApiKeyButton: document.getElementById('enterApiKeyButton'),
                saveApiKeyButton: document.getElementById('saveApiKeyButton'),
                cancelApiKeyButton: document.getElementById('cancelApiKeyButton'),
                showApiKeyCheckbox: document.getElementById('showApiKeyCheckbox'),
                insertButton: document.getElementById('insertButton'),
                insertSnippetButton: document.getElementById('insertSnippetButton'),
                discardButton: document.getElementById('discardButton'),
                enterModelButton: document.getElementById('enterModelButton'),
                saveModelButton: document.getElementById('saveModelButton'),
                cancelModelButton: document.getElementById('cancelModelButton'),
                configureChatButton: document.getElementById('configureChatButton'),
                saveChatConfigButton: document.getElementById('saveChatConfigButton'),
                cancelChatConfigButton: document.getElementById('cancelChatConfigButton'),
                apiKeyModal: document.getElementById('apiKeyModal'),
                modelModal: document.getElementById('modelModal'),
                chatConfigModal: document.getElementById('chatConfigModal'),
                changeStyleButton: document.getElementById('changeStyleButton'),
                styleModal: document.getElementById('styleModal'),
                saveStyleButton: document.getElementById('saveStyleButton'),
                cancelStyleButton: document.getElementById('cancelStyleButton'),
                modelInput: document.getElementById('modelInput'),
                styleModalInput: document.getElementById('styleModalInput'),
            };

            const missingElements = Object.entries(elements)
                .filter(([key, value]) => !value)
                .map(([key]) => key);
            if (missingElements.length > 0) {
                console.error('initializeEventHandlers: Missing DOM elements:', missingElements.join(', '));
                const statusMessage = document.getElementById('statusMessage');
                if (statusMessage) {
                    statusMessage.textContent = 'Error: Some UI elements are missing.';
                }
                return;
            }

            elements.submitButton.addEventListener('click', submitQuestion);
            elements.questionBox.addEventListener('keydown', handleQuestionBoxKeydown);
            elements.cancelButton.addEventListener('click', cancelQuestion);
            elements.discardButton.addEventListener('click', discardResponse);
            elements.insertButton.addEventListener('click', insertResponse);
            elements.insertSnippetButton.addEventListener('click', insertSnippet);
            elements.enterApiKeyButton.addEventListener('click', () => {
                elements.apiKeyModal.style.display = 'block';
                document.getElementById('apiKeyInput')?.focus();
            });
            elements.saveApiKeyButton.addEventListener('click', saveApiKey);
            elements.cancelApiKeyButton.addEventListener('click', () => {
                elements.apiKeyModal.style.display = 'none';
                const statusMessage = document.getElementById('statusMessage');
                if (statusMessage) {
                    statusMessage.textContent = isApiKeyValid ? 'API key unchanged.' : 'Please enter a valid API key.';
                }
            });
            elements.showApiKeyCheckbox.addEventListener('change', toggleApiKeyVisibility);
            elements.enterModelButton.addEventListener('click', () => {
                elements.modelModal.style.display = 'block';
                document.getElementById('modelInput')?.focus();
            });
            elements.saveModelButton.addEventListener('click', saveModel);
            elements.cancelModelButton.addEventListener('click', () => {
                elements.modelModal.style.display = 'none';
                const statusMessage = document.getElementById('statusMessage');
                if (statusMessage) {
                    statusMessage.textContent = currentModel ? 'Model unchanged.' : 'Please select a valid model.';
                }
            });
            elements.configureChatButton.addEventListener('click', configureChat);
            elements.saveChatConfigButton.addEventListener('click', saveChatConfig);
            elements.cancelChatConfigButton.addEventListener('click', cancelChatConfig);
            elements.changeStyleButton.addEventListener('click', () => {
                elements.styleModal.style.display = 'block';
                const styleModalInput = document.getElementById('styleModalInput');
                styleModalInput.value = currentStyle;
                styleModalInput.focus();
            });
            elements.saveStyleButton.addEventListener('click', saveStyle);
            elements.cancelStyleButton.addEventListener('click', cancelStyle);
            elements.modelInput.addEventListener('keydown', (event) => {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    saveModel();
                }
            });
            elements.styleModalInput.addEventListener('keydown', (event) => {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    saveStyle();
                }
            });

            console.log('initializeEventHandlers: Event handlers set up successfully');
        } catch (error) {
            console.error('initializeEventHandlers: Error setting up event handlers', error.message, error.stack);
            const statusMessage = document.getElementById('statusMessage');
            if (statusMessage) {
                statusMessage.textContent = `Error initializing UI: ${error.message}`;
            }
        }
    }

    Office.onReady(() => {
        console.log('Office.onReady: Office initialized');
        initializeEventHandlers();
        checkConfigOnLoad();
    });
}