// api/quota-check.js — Tests if API keys are active and responding
// Called by the quota dashboard tab in the frontend

// Known Gemini free-tier models to try in order
const GEMINI_MODELS = [
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite',
    'gemini-1.5-flash',
    'gemini-1.5-flash-latest',
    'gemini-1.5-flash-8b',
    'gemini-pro',
];

export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }

    const { api } = req.query;
    const start = Date.now();

    // Special endpoint: find working model for a key
    if (api === 'find-model') {
        const key = process.env.GEMINI_API_KEY;
        if (!key) return res.status(200).json({ ok: false, error: 'No key configured' });

        for (const model of GEMINI_MODELS) {
            const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
            try {
                const r = await fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        contents: [{ parts: [{ text: 'hi' }] }],
                        generationConfig: { maxOutputTokens: 5, temperature: 0 }
                    })
                });
                const body = await r.text();
                if (r.ok) {
                    return res.status(200).json({ ok: true, model, ms: Date.now() - start });
                }
                // If quota exhausted, still note it worked before
                if (r.status === 429) {
                    return res.status(200).json({ ok: false, exhausted: true, model, ms: Date.now() - start });
                }
            } catch (_) {}
        }
        return res.status(200).json({ ok: false, error: 'No working model found for this key' });
    }

    try {
        if (api === 'gemini' || api === 'gemini2') {
            const key = api === 'gemini2'
                ? process.env.GEMINI_API_KEY_2
                : process.env.GEMINI_API_KEY;

            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            // Try each model until one works
            for (const model of GEMINI_MODELS) {
                const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
                const r = await fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        contents: [{ parts: [{ text: 'Reply with just: {"ok":true}' }] }],
                        generationConfig: { maxOutputTokens: 20, temperature: 0 }
                    })
                });

                const ms = Date.now() - start;
                const body = await r.text();

                if (r.status === 429) {
                    return res.status(200).json({ ok: false, exhausted: true, ms, model });
                }

                if (r.ok) {
                    return res.status(200).json({ ok: true, ms, model });
                }

                // Model not found — try next
                if (r.status === 404 || body.includes('not found')) continue;

                // Other error
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) {}
                return res.status(200).json({ ok: false, error: errMsg, ms });
            }

            return res.status(200).json({ ok: false, error: 'No supported Gemini model found for this key' });

        } else if (api === 'groq') {
            const key = process.env.GROQ_API_KEY;
            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            const r = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
                body: JSON.stringify({
                    model: 'llama-3.1-8b-instant',
                    max_tokens: 5,
                    temperature: 0,
                    messages: [{ role: 'user', content: 'Say hi' }]
                })
            });

            const ms = Date.now() - start;
            const body = await r.text();

            if (r.status === 429) return res.status(200).json({ ok: false, exhausted: true, ms });
            if (!r.ok) {
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) {}
                return res.status(200).json({ ok: false, error: errMsg, ms });
            }
            return res.status(200).json({ ok: true, ms });

        } else {
            return res.status(400).json({ ok: false, error: 'Invalid api parameter' });
        }

    } catch (err) {
        return res.status(200).json({ ok: false, error: err.message, ms: Date.now() - start });
    }
}

    try {
        if (api === 'gemini' || api === 'gemini2') {
            const key = api === 'gemini2'
                ? process.env.GEMINI_API_KEY_2
                : process.env.GEMINI_API_KEY;

            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            // Minimal test prompt — 1 token in, 1 token out
            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-05-20:generateContent?key=${key}`;
            const r = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: 'Reply with just: {"ok":true}' }] }],
                    generationConfig: { maxOutputTokens: 20, temperature: 0 }
                })
            });

            const ms = Date.now() - start;
            const body = await r.text();

            if (r.status === 429) {
                return res.status(200).json({ ok: false, exhausted: true, ms });
            }
            if (!r.ok) {
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) {}
                return res.status(200).json({ ok: false, error: errMsg, ms });
            }

            return res.status(200).json({ ok: true, ms });

        } else if (api === 'groq') {
            const key = process.env.GROQ_API_KEY;
            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            const r = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
                body: JSON.stringify({
                    model: 'llama-3.1-8b-instant',
                    max_tokens: 5,
                    temperature: 0,
                    messages: [{ role: 'user', content: 'Say hi' }]
                })
            });

            const ms = Date.now() - start;
            const body = await r.text();

            if (r.status === 429) {
                return res.status(200).json({ ok: false, exhausted: true, ms });
            }
            if (!r.ok) {
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) {}
                return res.status(200).json({ ok: false, error: errMsg, ms });
            }

            return res.status(200).json({ ok: true, ms });

        } else {
            return res.status(400).json({ ok: false, error: 'Invalid api parameter' });
        }

    } catch (err) {
        return res.status(200).json({ ok: false, error: err.message, ms: Date.now() - start });
    }
}
