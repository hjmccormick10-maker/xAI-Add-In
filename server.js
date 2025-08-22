const express = require('express');
const cors = require('cors');
const axios = require('axios');
const path = require('path');
const fs = require('fs').promises;
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'dist')));

// Load or initialize config
const configPath = process.env.APPDATA
    ? path.join(process.env.APPDATA, 'Microsoft', 'AddIns', 'MinWordAdd', 'xAI_Config.json')
    : path.join(__dirname, 'xAI_Config.json');
let config = {
    XAI_API_KEY: '',
    XAI_MODEL: 'grok-3',
    XAI_MAX_TOKENS: 1000,
    XAI_SYSTEM_PROMPT: 'You are a concise assistant for proposal writing in a Word add-in.',
    XAI_FONT: 'Normal' // Added to support taskpane.js XAI_FONT handling
};

(async () => {
    console.log(`Loading config from: ${configPath}`);
    try {
        const exists = await fs.access(configPath).then(() => true).catch(() => false);
        if (exists) {
            const data = await fs.readFile(configPath, 'utf8');
            console.log(`Raw config content: ${data}`);
            const parsedConfig = JSON.parse(data);
            // Use XAI_API_KEY explicitly
            if (parsedConfig.XAI_API_KEY && /^[a-zA-Z0-9-]+$/.test(parsedConfig.XAI_API_KEY) && parsedConfig.XAI_API_KEY.startsWith('xai-') && parsedConfig.XAI_API_KEY.length >= 20) {
                config.XAI_API_KEY = parsedConfig.XAI_API_KEY;
                console.log(`Valid API key loaded (first 10 chars): ${parsedConfig.XAI_API_KEY.substring(0, 10)}`);
            } else {
                config.XAI_API_KEY = '';
                console.log(`Invalid or missing API key in config: ${parsedConfig.XAI_API_KEY || 'none'}`);
            }
            // Load XAI_MODEL without validation
            if (parsedConfig.XAI_MODEL) {
                config.XAI_MODEL = parsedConfig.XAI_MODEL;
            }
            if (parsedConfig.XAI_MAX_TOKENS && Number.isInteger(parsedConfig.XAI_MAX_TOKENS) && parsedConfig.XAI_MAX_TOKENS > 0) {
                config.XAI_MAX_TOKENS = parsedConfig.XAI_MAX_TOKENS;
            }
            if (typeof parsedConfig.XAI_SYSTEM_PROMPT === 'string') {
                config.XAI_SYSTEM_PROMPT = parsedConfig.XAI_SYSTEM_PROMPT;
            }
            // Load XAI_FONT
            if (typeof parsedConfig.XAI_FONT === 'string') {
                config.XAI_FONT = parsedConfig.XAI_FONT;
                console.log(`XAI_FONT loaded: ${parsedConfig.XAI_FONT}`);
            }
            console.log('Config loaded:', { ...config, XAI_API_KEY: '[REDACTED]' });
            // Clean config file to remove lowercase fields
            const cleanedConfig = {
                XAI_API_KEY: config.XAI_API_KEY,
                XAI_MODEL: config.XAI_MODEL,
                XAI_MAX_TOKENS: config.XAI_MAX_TOKENS,
                XAI_SYSTEM_PROMPT: config.XAI_SYSTEM_PROMPT,
                XAI_FONT: config.XAI_FONT
            };
            await fs.writeFile(configPath, JSON.stringify(cleanedConfig, null, 2));
            console.log('Config file cleaned, lowercase fields removed');
        } else {
            console.log('No config file found, creating default');
            await fs.mkdir(path.dirname(configPath), { recursive: true });
            await fs.writeFile(configPath, JSON.stringify(config, null, 2));
        }
    } catch (error) {
        console.error('Error loading config:', error.message);
        config.XAI_API_KEY = '';
        config.XAI_FONT = 'Normal'; // Ensure default if config load fails
        await fs.mkdir(path.dirname(configPath), { recursive: true });
        await fs.writeFile(configPath, JSON.stringify(config, null, 2)).catch(err => console.error('Failed to create default config:', err.message));
    }
})();

// Temporary debug endpoint to verify loaded API key (REMOVE AFTER VERIFICATION)
app.get('/debug-api-key', (req, res) => {
    console.log('GET /debug-api-key called');
    res.json({ loadedApiKey: config.XAI_API_KEY || 'none' });
});

// Validate API key
app.post('/validate', async (req, res) => {
    console.log('POST /validate called:', { ...req.body, apiKey: '[REDACTED]' });
    try {
        const { apiKey } = req.body;
        if (!apiKey || !apiKey.startsWith('xai-') || apiKey.length < 20 || !/^[a-zA-Z0-9-]+$/.test(apiKey)) {
            return res.status(400).json({ valid: false, error: 'Invalid API key format. Must start with "xai-" and be at least 20 characters.' });
        }
        const response = await axios.post('https://api.x.ai/v1/chat/completions', {
            messages: [{ role: 'user', content: 'Test' }],
            model: 'grok-3',
            max_tokens: 10
        }, {
            headers: { Authorization: `Bearer ${apiKey}` },
            timeout: 10000
        });
        if (response.status === 200) {
            console.log('Validate: API key valid');
            return res.json({ valid: true });
        }
        return res.status(401).json({ valid: false, error: 'API key authentication failed' });
    } catch (error) {
        console.error('Validate error:', error.message);
        const status = error.response?.status || 500;
        const errorMessage = error.response?.data?.error || error.message || 'Failed to validate API key';
        return res.status(status).json({ valid: false, error: errorMessage });
    }
});

// Validate model
app.post('/validate-model', async (req, res) => {
    console.log('POST /validate-model called:', req.body);
    try {
        const { model } = req.body;
        if (!model) {
            return res.status(400).json({ valid: false, error: 'Model name is required' });
        }
        if (!config.XAI_API_KEY) {
            return res.status(401).json({ valid: false, error: 'No API key configured' });
        }
        const response = await axios.post('https://api.x.ai/v1/chat/completions', {
            messages: [{ role: 'user', content: 'Test' }],
            model: model,
            max_tokens: 10
        }, {
            headers: { Authorization: `Bearer ${config.XAI_API_KEY}` },
            timeout: 10000
        });
        if (response.status === 200) {
            console.log('Validate-model: Model valid');
            return res.json({ valid: true });
        }
        return res.status(400).json({ valid: false, error: 'Invalid model' });
    } catch (error) {
        console.error('Validate-model error:', error.message);
        const status = error.response?.status || 500;
        const errorMessage = error.response?.data?.error || error.message || 'Failed to validate model';
        return res.status(status).json({ valid: false, error: errorMessage });
    }
});

// Get config
app.get('/get-config', (req, res) => {
    console.log('GET /get-config called');
    const hasValidApiKey = config.XAI_API_KEY && config.XAI_API_KEY.startsWith('xai-') && config.XAI_API_KEY.length >= 20 && /^[a-zA-Z0-9-]+$/.test(config.XAI_API_KEY);
    res.json({
        XAI_API_KEY: hasValidApiKey ? '[REDACTED]' : '',
        XAI_MODEL: config.XAI_MODEL,
        XAI_MAX_TOKENS: config.XAI_MAX_TOKENS,
        XAI_SYSTEM_PROMPT: config.XAI_SYSTEM_PROMPT,
        XAI_FONT: config.XAI_FONT, // Added to include XAI_FONT
        hasValidApiKey
    });
});

// Save config
app.post('/save-env', async (req, res) => {
    console.log('POST /save-env called:', { ...req.body, apiKey: req.body.apiKey ? '[REDACTED]' : undefined, font: req.body.font });
    try {
        const { apiKey, model, font } = req.body; // Client sends 'apiKey', 'model', and/or 'font'
        const newConfig = { ...config };
        if (apiKey) {
            newConfig.XAI_API_KEY = apiKey;
        }
        if (model) {
            newConfig.XAI_MODEL = model;
        }
        if (font) {
            newConfig.XAI_FONT = font;
        }
        // Save only uppercase fields
        const cleanedConfig = {
            XAI_API_KEY: newConfig.XAI_API_KEY,
            XAI_MODEL: newConfig.XAI_MODEL,
            XAI_MAX_TOKENS: newConfig.XAI_MAX_TOKENS,
            XAI_SYSTEM_PROMPT: newConfig.XAI_SYSTEM_PROMPT,
            XAI_FONT: newConfig.XAI_FONT
        };
        await fs.mkdir(path.dirname(configPath), { recursive: true });
        await fs.writeFile(configPath, JSON.stringify(cleanedConfig, null, 2));
        config = newConfig;
        console.log('Config saved:', { ...config, XAI_API_KEY: '[REDACTED]' });
        res.json({ success: true });
    } catch (error) {
        console.error('Error saving config:', error.message);
        res.status(500).json({ error: `Failed to save config: ${error.message}` });
    }
});

// API chat endpoint
app.post('/api/chat', async (req, res) => {
    console.log('POST /api/chat called:', { ...req.body, messages: req.body.messages.map(msg => ({ ...msg, content: '[TRUNCATED]' })) });
    try {
        const { messages, model, max_tokens } = req.body;
        if (!config.XAI_API_KEY) {
            return res.status(401).json({ error: 'No API key configured' });
        }
        const effectiveModel = model || config.XAI_MODEL;
        const effectiveMaxTokens = max_tokens || config.XAI_MAX_TOKENS;
        const systemMessage = { role: 'system', content: config.XAI_SYSTEM_PROMPT };
        const requestMessages = [systemMessage, ...messages];
        const response = await axios.post('https://api.x.ai/v1/chat/completions', {
            messages: requestMessages,
            model: effectiveModel,
            max_tokens: effectiveMaxTokens,
            stream: false
        }, {
            headers: { Authorization: `Bearer ${config.XAI_API_KEY}` },
            timeout: 30000
        });
        res.json(response.data);
    } catch (error) {
        console.error('API error:', error.message);
        const status = error.response?.status || 500;
        const errorMessage = error.response?.data?.error || error.message || 'Failed to process chat request';
        res.status(status).json({ error: errorMessage });
    }
});

app.listen(3001, () => console.log('Server running on port 3001'));