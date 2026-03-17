// api/generate.js — Vercel Serverless Function
// Primary:  Gemini 1.5 Flash (Google) — 1M token/day free
// Fallback: Groq Llama 3.3 70B       — 500K token/day free
// DOCX generation: docx npm package

import {
    Document, Packer, Paragraph, TextRun,
    AlignmentType, BorderStyle, PageNumber,
    Header, Footer, ShadingType, ImageRun,
    TabStopType, TabStopPosition
} from 'docx';

// XtremeLabs logo (base64 PNG — embedded at build time)
const LOGO_B64 = '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';

// Increase Vercel body size limit for large file uploads (25 md files)
export const config = {
    api: {
        bodyParser: {
            sizeLimit: '10mb'
        }
    }
};

export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }
    if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

    try {
        const { files, courseName } = req.body;

        if (!files || !Array.isArray(files) || files.length === 0)
            return res.status(400).json({ error: 'files array is required' });
        if (!courseName || typeof courseName !== 'string')
            return res.status(400).json({ error: 'courseName is required' });

        const geminiKey = process.env.GEMINI_API_KEY;
        const groqKey   = process.env.GROQ_API_KEY;

        if (!geminiKey && !groqKey)
            return res.status(500).json({ error: 'No API keys configured on server' });

        const combinedContent = files.map((f, i) =>
            `=== LAB FILE ${i + 1}: ${f.name} ===\n${f.content}`
        ).join('\n\n');

        let summaryData = null;
        let usedApi = '';

        if (geminiKey) {
            try {
                // Gemini: 1M context — use up to 200K chars
                const prompt = buildPrompt(files.length, courseName, combinedContent, 200000);
                summaryData = await callGemini(geminiKey, prompt);
                usedApi = 'Gemini';
            } catch (geminiErr) {
                console.error('Gemini failed:', geminiErr.message, '— trying Groq fallback');
                if (!groqKey) throw new Error(`Gemini failed: ${geminiErr.message}. No Groq fallback configured.`);
            }
        }

        if (!summaryData && groqKey) {
            try {
                // Groq llama3-70b-8192: 8K context — trim to 20K chars (~6K tokens for content)
                const prompt = buildPrompt(files.length, courseName, combinedContent, 20000);
                summaryData = await callGroq(groqKey, prompt);
                usedApi = 'Groq';
            } catch (groqErr) {
                throw new Error(`Both APIs failed. Groq error: ${groqErr.message}`);
            }
        }

        if (!summaryData) throw new Error('Failed to generate summary — no API available');

        const docBuffer = await buildDocx(summaryData, courseName);
        const base64 = docBuffer.toString('base64');
        const filename = `${courseName.replace(/[^a-z0-9]/gi, '_')}_Summary.docx`;

        return res.status(200).json({ ok: true, docx: base64, filename, summary: summaryData, usedApi });

    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

// ═══════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════
function parseJsonSafe(raw) {
    if (!raw || raw.trim() === '') throw new Error('Empty response from AI');
    // Strip markdown fences if present
    let cleaned = raw
        .replace(/^```json\s*/i, '')
        .replace(/^```\s*/i, '')
        .replace(/```\s*$/i, '')
        .trim();
    // Find first { and last } to extract JSON block
    const start = cleaned.indexOf('{');
    const end   = cleaned.lastIndexOf('}');
    if (start === -1 || end === -1) throw new Error('No JSON object found in AI response');
    cleaned = cleaned.substring(start, end + 1);
    return JSON.parse(cleaned);
}

// ═══════════════════════════════════════════════
// GEMINI API CALL
// ═══════════════════════════════════════════════
async function callGemini(apiKey, prompt) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

    const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
                temperature: 0.4,
                maxOutputTokens: 8192
                // Note: responseMimeType removed — causes issues with some accounts
            },
            systemInstruction: {
                parts: [{ text: 'You are a professional technical writer. Always respond with valid JSON only — no markdown fences, no explanation, no extra text.' }]
            }
        })
    });

    // Read response as text first to avoid "Unexpected end of JSON" errors
    const rawBody = await res.text();

    if (!res.ok) {
        let errMsg = `Gemini API error ${res.status}`;
        try {
            const errJson = JSON.parse(rawBody);
            errMsg = errJson.error?.message || errMsg;
        } catch (_) {}
        if (res.status === 429 || errMsg.includes('quota') || errMsg.includes('RESOURCE_EXHAUSTED'))
            throw new Error(`Gemini quota exceeded: ${errMsg}`);
        throw new Error(errMsg);
    }

    const data = JSON.parse(rawBody);

    // Check for blocked content
    const candidate = data.candidates?.[0];
    if (!candidate) throw new Error('Gemini returned no candidates');
    if (candidate.finishReason === 'SAFETY') throw new Error('Gemini blocked response for safety reasons');

    const rawText = candidate.content?.parts?.[0]?.text || '';
    return parseJsonSafe(rawText);
}

// ═══════════════════════════════════════════════
// GROQ API CALL
// Using llama-3.1-8b-instant — 131K context, 6000 TPM free
// ═══════════════════════════════════════════════
async function callGroq(apiKey, prompt) {
    const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: 'llama-3.1-8b-instant',
            max_tokens: 8000,
            temperature: 0.4,
            messages: [
                {
                    role: 'system',
                    content: 'You are a professional technical writer. Always respond with valid JSON only — no markdown fences, no explanation, no extra text.'
                },
                { role: 'user', content: prompt }
            ]
        })
    });

    const rawBody = await res.text();

    if (!res.ok) {
        let errMsg = `Groq API error ${res.status}`;
        try { errMsg = JSON.parse(rawBody).error?.message || errMsg; } catch (_) {}
        throw new Error(errMsg);
    }

    const data = JSON.parse(rawBody);
    const rawText = data.choices?.[0]?.message?.content || '';
    return parseJsonSafe(rawText);
}

// ═══════════════════════════════════════════════
// PROMPT BUILDER
// Gemini: 200K chars limit (1M token context)
// Groq:   20K chars limit (8K token context)
// ═══════════════════════════════════════════════
function buildPrompt(fileCount, courseName, content, maxChars = 200000) {
    const trimmed = content.length > maxChars
        ? content.substring(0, maxChars) + '\n\n...[remaining files truncated due to length]'
        : content;

    return `I have ${fileCount} Markdown files containing labs for the "${courseName}" course. I need you to analyze all of them in detail and generate a professional summary document for the course.

The output should follow this structure:

Course Overview (at the top):
Start with a short 3-5 sentence summary of the entire course. This should capture the overall purpose of the course, the types of skills being developed, and the areas covered.

Detailed Lab Summaries (one per lab):
Each lab should be clearly titled (e.g., Lab 1: Understanding Threats and Vulnerabilities) and followed by four points in the following format:
- Objective: A one-sentence description of the lab's goal.
- Key Topics Covered: A brief mention (1-2 lines) of core concepts or technologies discussed.
- Hands-On Activity: A short summary of the tasks or configurations performed in the lab.
- Real-World Application: A sentence on how this lab's skills or concepts apply in a real IT/security environment.

Use a professional tone and ensure clarity, accuracy, and consistency across all lab entries.

Reference example:
Lab 4: Enforcing Password Security Policies
Objective: To understand and configure password policies at local and domain levels.
Key Topics Covered: Local security policy, domain security policy, and user authentication.
Hands-On Activity: Removing a Windows computer from a domain, modifying local security policies, creating a local user for testing, allowing logons via remote desktop, verifying local and domain password policies, modifying domain security policies, and unlocking user accounts.
Real-World Application: IT professionals use these skills to manage user authentication and enhance system security by enforcing strong password policies.

Respond with ONLY this JSON structure (no markdown fences, no extra text):
{
  "courseOverview": "3-5 sentence summary here.",
  "labs": [
    {
      "title": "Lab 1: Full Lab Title Here",
      "objective": "One-sentence description of the lab goal.",
      "keyTopics": "1-2 lines on core concepts or technologies discussed.",
      "handsOnActivity": "Short summary of tasks or configurations performed.",
      "realWorldApplication": "One sentence on real-world IT application."
    }
  ]
}

LAB FILES:
${trimmed}`;
}

// ═══════════════════════════════════════════════
// DOCX BUILDER
// - Course title: Aptos 16pt bold centered, dark blue
// - Section headings: Times New Roman 14pt bold, dark blue
// - Lab titles + sub-headings: Times New Roman 12pt bold
// - Body: Times New Roman 12pt
// - Header: course name left + logo right
// - Footer: "Powered By:" bold + logo centered
// - Page border: professional double-line design
// ═══════════════════════════════════════════════
async function buildDocx(data, courseName) {
    const DARK_BLUE      = '1F3864';
    const ACCENT         = '2E5FA3';
    const DARK_GREY      = '2F2F2F';
    const MID_GREY       = 'CCCCCC';
    // Light bg shading colors
    const BG_SECTION     = 'D6E4F0';   // light blue for Course Overview / Detailed Lab Summaries
    const BG_LAB         = 'EBF3FB';   // very light blue for lab title callout

    const logoBuffer = Buffer.from(LOGO_B64, 'base64');
    const children = [];

    // ── Course Title ──
    children.push(
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 320 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: ACCENT } },
            children: [
                new TextRun({ text: courseName, bold: true, size: 32, font: 'Aptos', color: DARK_BLUE })
            ]
        })
    );

    // ── Course Overview heading — light bg shading ──
    children.push(
        new Paragraph({
            spacing: { before: 240, after: 120 },
            shading: { type: ShadingType.CLEAR, fill: BG_SECTION },
            indent: { left: 120, right: 120 },
            children: [
                new TextRun({ text: 'Course Overview', bold: true, size: 28, font: 'Times New Roman', color: DARK_BLUE })
            ]
        }),
        new Paragraph({
            spacing: { before: 120, after: 360 },
            children: [
                new TextRun({ text: data.courseOverview || '', size: 24, font: 'Times New Roman', color: DARK_GREY })
            ]
        })
    );

    // ── Detailed Lab Summaries heading — light bg shading ──
    children.push(
        new Paragraph({
            spacing: { before: 160, after: 200 },
            shading: { type: ShadingType.CLEAR, fill: BG_SECTION },
            indent: { left: 120, right: 120 },
            children: [
                new TextRun({ text: 'Detailed Lab Summaries', bold: true, size: 28, font: 'Times New Roman', color: DARK_BLUE })
            ]
        })
    );

    // ── Lab Entries ──
    for (let i = 0; i < data.labs.length; i++) {
        const lab = data.labs[i];

        // Lab title callout — light bg + thick left border
        children.push(
            new Paragraph({
                spacing: { before: 260, after: 100 },
                shading: { type: ShadingType.CLEAR, fill: BG_LAB },
                border: {
                    left: { style: BorderStyle.THICK, size: 20, color: ACCENT }
                },
                indent: { left: 200, right: 120 },
                children: [
                    new TextRun({ text: lab.title, bold: true, size: 24, font: 'Times New Roman', color: DARK_BLUE })
                ]
            })
        );

        const sections = [
            { label: 'Objective:',              value: lab.objective },
            { label: 'Key Topics Covered:',     value: lab.keyTopics },
            { label: 'Hands-On Activity:',      value: lab.handsOnActivity },
            { label: 'Real-World Application:', value: lab.realWorldApplication }
        ];

        for (let j = 0; j < sections.length; j++) {
            const s = sections[j];
            const isLast = j === sections.length - 1;
            children.push(
                new Paragraph({
                    spacing: { before: 80, after: isLast ? 140 : 60 },
                    indent: { left: 400 },
                    children: [
                        new TextRun({ text: s.label + ' ', bold: true, size: 24, font: 'Times New Roman', color: DARK_GREY }),
                        new TextRun({ text: s.value || '', size: 24, font: 'Times New Roman', color: DARK_GREY })
                    ]
                })
            );
        }

        if (i < data.labs.length - 1) {
            children.push(
                new Paragraph({
                    spacing: { before: 60, after: 60 },
                    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: MID_GREY } },
                    children: []
                })
            );
        }
    }

    // ── Header: logo on the right with small right indent to sit near page edge ──
    const headerParagraph = new Paragraph({
        spacing: { before: 0, after: 0, line: 240 },
        tabStops: [{ type: TabStopType.RIGHT, position: 9260 }],
        children: [
            new TextRun({ text: '\t' }),
            new ImageRun({
                data: logoBuffer,
                type: 'png',
                transformation: { width: 120, height: 27 }
            })
        ]
    });

    // ── Footer: "Powered By:" + logo — use vertical alignment to fix logo rising ──
    const footerParagraph = new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT } },
        spacing: { before: 80, after: 0, line: 360, lineRule: 'auto' },
        children: [
            new TextRun({
                text: 'Powered By:  ',
                bold: true,
                size: 20,
                font: 'Times New Roman',
                color: DARK_BLUE
            }),
            new ImageRun({
                data: logoBuffer,
                type: 'png',
                transformation: { width: 95, height: 21 },
            })
        ]
    });

    // ── Page border ──
    const pageBorder = {
        pageBorderTop:    { style: BorderStyle.TRIPLE, size: 24, color: DARK_BLUE, space: 1 },
        pageBorderBottom: { style: BorderStyle.TRIPLE, size: 24, color: DARK_BLUE, space: 1 },
        pageBorderLeft:   { style: BorderStyle.TRIPLE, size: 24, color: DARK_BLUE, space: 1 },
        pageBorderRight:  { style: BorderStyle.TRIPLE, size: 24, color: DARK_BLUE, space: 1 }
    };

    const doc = new Document({
        styles: {
            default: { document: { run: { font: 'Times New Roman', size: 24, color: DARK_GREY } } }
        },
        sections: [{
            properties: {
                page: {
                    size: { width: 12240, height: 15840 },
                    margin: { top: 1080, right: 1080, bottom: 1200, left: 1080 },
                    borders: pageBorder
                }
            },
            headers: { default: new Header({ children: [headerParagraph] }) },
            footers: { default: new Footer({ children: [footerParagraph] }) },
            children
        }]
    });

    // ── Patch XML: offsetFrom="page" + fix logo vertical alignment in footer ──
    const rawBuffer = await Packer.toBuffer(doc);
    return patchDocxXml(rawBuffer);
}

// Post-process docx XML:
// 1. offsetFrom="page" — border hugs page edge
// 2. Fix logo vertical baseline in footer (distT=0 so image sits on text baseline)
async function patchDocxXml(buffer) {
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);

    // Patch document.xml — page border offsetFrom
    const docXmlPath = 'word/document.xml';
    let docXml = await zip.file(docXmlPath).async('string');
    docXml = docXml.replace(
        /<w:pgBorders(?![^>]*w:offsetFrom)/g,
        '<w:pgBorders w:offsetFrom="page"'
    );
    zip.file(docXmlPath, docXml);

    // Patch footer XML — fix logo vertical position (distT/distB to 0)
    const footerFiles = Object.keys(zip.files).filter(f => f.startsWith('word/footer'));
    for (const fPath of footerFiles) {
        let fXml = await zip.file(fPath).async('string');
        // Set distT="0" distB="0" on inline drawing so image aligns with text baseline
        fXml = fXml.replace(
            /<wp:inline([^>]*)>/g,
            (match, attrs) => {
                let updated = attrs
                    .replace(/distT="[^"]*"/, 'distT="0"')
                    .replace(/distB="[^"]*"/, 'distB="0"');
                if (!updated.includes('distT=')) updated += ' distT="0"';
                if (!updated.includes('distB=')) updated += ' distB="0"';
                return `<wp:inline${updated}>`;
            }
        );
        zip.file(fPath, fXml);
    }

    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
