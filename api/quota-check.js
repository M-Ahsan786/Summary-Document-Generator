// api/quota-check.js

const GEMINI_MODELS = [
    'gemini-3-flash-preview',
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite',
    'gemini-1.5-flash',
    'gemini-1.5-flash-latest',
];

export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }

    const { api } = req.query;
    const start = Date.now();

    try {
        if (api === 'gemini' || api === 'gemini2') {
            const key = api === 'gemini2'
                ? process.env.GEMINI_API_KEY_2
                : process.env.GEMINI_API_KEY;

            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            for (const model of GEMINI_MODELS) {
                const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
                let r, body;
                try {
                    r = await fetch(url, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            contents: [{ parts: [{ text: 'hi' }] }],
                            generationConfig: { maxOutputTokens: 5, temperature: 0 }
                        })
                    });
                    body = await r.text();
                } catch (fetchErr) {
                    continue;
                }

                const ms = Date.now() - start;

                // Quota exhausted
                if (r.status === 429) {
                    return res.status(200).json({ ok: false, exhausted: true, ms, model });
                }

                // Model not found — try next
                if (r.status === 404 || body.includes('not found for API version') || body.includes('is not supported')) {
                    continue;
                }

                // Success
                if (r.ok) {
                    return res.status(200).json({ ok: true, ms, model });
                }

                // Other error — parse and return
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) { errMsg = body.substring(0, 100); }

                // If it's a quota/auth error stop trying
                if (r.status === 403 || errMsg.includes('quota') || errMsg.includes('API_KEY_INVALID')) {
                    return res.status(200).json({ ok: false, exhausted: true, ms, error: errMsg });
                }

                // Otherwise try next model
                continue;
            }

            return res.status(200).json({ ok: false, error: 'No supported Gemini model found for this API key', ms: Date.now() - start });

        } else if (api === 'groq') {
            const key = process.env.GROQ_API_KEY;
            if (!key) return res.status(200).json({ ok: false, error: 'Key not configured in environment variables' });

            let r, body;
            try {
                r = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
                    body: JSON.stringify({
                        model: 'llama-3.1-8b-instant',
                        max_tokens: 5, temperature: 0,
                        messages: [{ role: 'user', content: 'hi' }]
                    })
                });
                body = await r.text();
            } catch (fetchErr) {
                return res.status(200).json({ ok: false, error: fetchErr.message, ms: Date.now() - start });
            }

            const ms = Date.now() - start;
            if (r.status === 429) return res.status(200).json({ ok: false, exhausted: true, ms });
            if (!r.ok) {
                let errMsg = `HTTP ${r.status}`;
                try { errMsg = JSON.parse(body).error?.message || errMsg; } catch (_) {}
                return res.status(200).json({ ok: false, error: errMsg, ms });
            }
            return res.status(200).json({ ok: true, ms });

        } else {
            return res.status(200).json({ ok: false, error: 'Invalid api parameter' });
        }

    } catch (err) {
        return res.status(200).json({ ok: false, error: err.message, ms: Date.now() - start });
    }
}
