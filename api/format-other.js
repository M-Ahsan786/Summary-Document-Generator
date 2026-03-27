// api/format-other.js — Vercel Serverless Function
// "Other Document" mode: preserves ALL original content.
// Only adds header logo, footer logo, page border via XML surgery.
//
// CRITICAL RULE: Image relationships for header/footer images MUST be in
// headerN.xml.rels / footerN.xml.rels — NOT in document.xml.rels.
// Word resolves images inside headers/footers from the header/footer's OWN .rels file.

import JSZip from 'jszip';

const LOGO_B64 = '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';

export const config = {
    api: { bodyParser: { sizeLimit: '20mb' } }
};

export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') { res.status(200).end(); return; }
    if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

    try {
        const { docxBase64, filename, options = {} } = req.body;
        if (!docxBase64) return res.status(400).json({ error: 'docxBase64 is required' });

        const opts = {
            headerLogo: options.headerLogo === true,
            footerLogo: options.footerLogo === true,
            pageBorder: options.pageBorder === true,
        };

        const inputBuffer = Buffer.from(docxBase64, 'base64');
        const outputBuffer = await applyFormatting(inputBuffer, opts);
        const outputBase64 = outputBuffer.toString('base64');
        const baseName = (filename || 'document').replace(/\.docx$/i, '');

        return res.status(200).json({
            ok: true,
            docx: outputBase64,
            filename: `${baseName}_Formatted.docx`,
            preview: { headerLogo: opts.headerLogo, footerLogo: opts.footerLogo, pageBorder: opts.pageBorder }
        });
    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

// ═══════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════
async function applyFormatting(buffer, opts) {
    const zip = await JSZip.loadAsync(buffer);
    const logoBuffer = Buffer.from(LOGO_B64, 'base64');

    if (opts.pageBorder) await applyPageBorder(zip);
    if (opts.headerLogo) await injectHeader(zip, logoBuffer);
    if (opts.footerLogo) await injectFooter(zip, logoBuffer);

    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ─────────────────────────────────────────────────────
// PAGE BORDER
// ─────────────────────────────────────────────────────
async function applyPageBorder(zip) {
    let xml = await zip.file('word/document.xml').async('string');
    xml = xml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/g, '');
    const border =
        `<w:pgBorders w:offsetFrom="page">` +
        `<w:top w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:left w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:bottom w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:right w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `</w:pgBorders>`;
    if (xml.includes('</w:sectPr>')) {
        xml = xml.replace(/<\/w:sectPr>/g, border + '</w:sectPr>');
    } else {
        xml = xml.replace(/<\/w:body>/, `<w:sectPr>${border}</w:sectPr></w:body>`);
    }
    zip.file('word/document.xml', xml);
}

// ─────────────────────────────────────────────────────
// HEADER LOGO
// ─────────────────────────────────────────────────────
async function injectHeader(zip, logoBuffer) {
    const imgFile = 'word/media/xtremelabs_logo_h.jpg';
    zip.file(imgFile, logoBuffer);
    await ensureContentType(zip, 'jpg', 'image/jpeg');

    // Find the single target header: the one wired as w:type="default" in document.xml
    // If none found, fall back to first existing header, or create a new one
    const targetHeader = await findOrCreateDefaultHeader(zip, 'header');

    const relsPath = `word/_rels/${targetHeader.replace('word/', '')}.rels`;
    let hXml = await zip.file(targetHeader).async('string');

    // Skip if our logo was already injected (idempotency)
    if (hXml.includes('xtremelabs_logo_h.jpg') || hXml.includes('xtremelabs_logo')) {
        return;
    }

    // Add image relationship to THIS header's own .rels file
    const imgRId = await addRelToFile(zip, relsPath, '../media/xtremelabs_logo_h.jpg',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');

    // Get unique drawing ID
    const drawingId = await getNextDrawingId(zip);

    // Inject logo paragraph at top of header
    hXml = ensureNs(hXml, 'w:hdr');
    hXml = hXml.replace(/(<w:hdr[^>]*>)/, `$1${logoParaHeader(imgRId, drawingId)}`);
    zip.file(targetHeader, hXml);
}

// ─────────────────────────────────────────────────────
// FOOTER LOGO
// ─────────────────────────────────────────────────────
async function injectFooter(zip, logoBuffer) {
    const imgFile = 'word/media/xtremelabs_logo_f.jpg';
    zip.file(imgFile, logoBuffer);
    await ensureContentType(zip, 'jpg', 'image/jpeg');

    // Find the single target footer: the one wired as w:type="default" in document.xml
    const targetFooter = await findOrCreateDefaultHeader(zip, 'footer');

    const relsPath = `word/_rels/${targetFooter.replace('word/', '')}.rels`;
    let fXml = await zip.file(targetFooter).async('string');

    // Skip if our logo was already injected (idempotency)
    if (fXml.includes('xtremelabs_logo_f.jpg') || fXml.includes('xtremelabs_logo')) {
        return;
    }

    // Add image relationship to THIS footer's own .rels file
    const imgRId = await addRelToFile(zip, relsPath, '../media/xtremelabs_logo_f.jpg',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');

    // Get unique drawing ID
    const drawingId = await getNextDrawingId(zip);

    // Inject logo paragraph at end of footer
    fXml = ensureNs(fXml, 'w:ftr');
    fXml = fXml.replace(/<\/w:ftr>/, `${logoParaFooter(imgRId, drawingId)}</w:ftr>`);
    zip.file(targetFooter, fXml);
}

// ─────────────────────────────────────────────────────
// Find (or create) the default header/footer xml path
// Returns path like 'word/header1.xml'
// ─────────────────────────────────────────────────────
async function findOrCreateDefaultHeader(zip, kind) {
    const tag = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
    const filePattern = kind === 'header' ? /^\/word\/header\d*\.xml$/ : /^\/word\/footer\d*\.xml$/;
    const filePatternNoSlash = kind === 'header' ? /^word\/header\d*\.xml$/ : /^word\/footer\d*\.xml$/;

    // Step 1: Find which headerN/footerN is wired as w:type="default" in document.xml
    const docXml = await zip.file('word/document.xml').async('string');
    const docRelsXml = zip.file('word/_rels/document.xml.rels')
        ? await zip.file('word/_rels/document.xml.rels').async('string')
        : '';

    // Find all default references: <w:headerReference w:type="default" r:id="rIdXXX"/>
    const defaultRefMatch = docXml.match(new RegExp(`<${tag}[^>]*w:type="default"[^>]*r:id="([^"]+)"`))
        || docXml.match(new RegExp(`<${tag}[^>]*r:id="([^"]+)"[^>]*w:type="default"`));

    if (defaultRefMatch) {
        const rId = defaultRefMatch[1];
        // Look up this rId in document.xml.rels to get the file name
        const relMatch = docRelsXml.match(new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`));
        if (relMatch) {
            const target = relMatch[1]; // e.g. "header1.xml" or "../word/header1.xml"
            const fileName = target.replace(/^.*\//, ''); // just "header1.xml"
            const fullPath = `word/${fileName}`;
            if (zip.file(fullPath)) return fullPath;
        }
    }

    // Step 2: Fall back to first existing header/footer file
    const existing = Object.keys(zip.files).filter(f => filePatternNoSlash.test(f));
    if (existing.length > 0) return existing[0];

    // Step 3: Create from scratch
    const newFile = kind === 'header' ? 'word/header1.xml' : 'word/footer1.xml';
    const relsFile = kind === 'header' ? 'word/_rels/header1.xml.rels' : 'word/_rels/footer1.xml.rels';
    const contentType = kind === 'header'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
    const relType = kind === 'header'
        ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
        : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
    const wiredRId = kind === 'header' ? 'rId_hdr1' : 'rId_ftr1';
    const emptyXml = kind === 'header'
        ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p></w:hdr>`
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:ftr>`;

    zip.file(newFile, emptyXml);
    zip.file(relsFile, relsDoc(''));
    await ensureOverride(zip, `/${newFile}`, contentType);
    await addRelToDocRels(zip, newFile.replace('word/', ''), relType, wiredRId);
    await wireRef(zip, kind, wiredRId);
    return newFile;
}

// ─────────────────────────────────────────────────────
// XML builders
// ─────────────────────────────────────────────────────
function relsDoc(relXml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        relXml +
        `</Relationships>`;
}

function buildHeader(rId, drawingId) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:hdr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
        ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
        ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
        ` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
        ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        logoParaHeader(rId, drawingId) + `</w:hdr>`;
}

function buildFooter(rId, drawingId) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:ftr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
        ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
        ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
        ` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"` +
        ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        logoParaFooter(rId, drawingId) + `</w:ftr>`;
}

function logoParaHeader(rId, drawingId) {
    return `<w:p><w:pPr><w:jc w:val="right"/></w:pPr>` +
        `<w:r>${imgDrawing(rId, 1270000, 292100, drawingId)}</w:r></w:p>`;
}

function logoParaFooter(rId, drawingId) {
    return `<w:p>` +
        `<w:pPr><w:jc w:val="center"/>` +
        `<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="2E5FA3"/></w:pBdr>` +
        `</w:pPr>` +
        `<w:r><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="20"/></w:rPr>` +
        `<w:t xml:space="preserve">Powered By:  </w:t></w:r>` +
        `<w:r>${imgDrawing(rId, 857250, 190500, drawingId)}</w:r></w:p>`;
}

function imgDrawing(rId, cx, cy, drawingId) {
    const did = drawingId || 1;
    return `<w:drawing>` +
        `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
        `<wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="${did}" name="logo_${did}"/><wp:cNvGraphicFramePr/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic><pic:nvPicPr><pic:cNvPr id="${did}" name="logo_${did}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rId}"/>` +
        `<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

// Get the next unique drawing ID across the entire docx (document + all headers/footers)
async function getNextDrawingId(zip) {
    const xmlFiles = Object.keys(zip.files).filter(f =>
        /^word\/(document|header\d*|footer\d*)\.xml$/.test(f)
    );
    let maxId = 0;
    for (const f of xmlFiles) {
        const xml = await zip.file(f).async('string');
        const matches = [...xml.matchAll(/\bdocPr\s+id="(\d+)"/g)];
        for (const m of matches) {
            const n = parseInt(m[1]);
            if (n > maxId) maxId = n;
        }
    }
    return maxId + 1;
}

// Ensure required namespaces on root element
function ensureNs(xml, tag) {
    const ns = {
        'xmlns:r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'xmlns:a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    };
    return xml.replace(new RegExp(`(<${tag}[^>]*)(>)`), (m, attrs, close) => {
        for (const [k, v] of Object.entries(ns)) {
            if (!attrs.includes(k)) attrs += ` ${k}="${v}"`;
        }
        return attrs + close;
    });
}

// ─────────────────────────────────────────────────────
// Relationship helpers
// ─────────────────────────────────────────────────────

// Set (or add) a specific rId → target in a .rels file
async function setRelInFile(zip, relsPath, rId, type, target) {
    let xml = zip.file(relsPath)
        ? await zip.file(relsPath).async('string')
        : relsDoc('');

    // Remove any existing entry with this rId
    xml = xml.replace(new RegExp(`<Relationship[^>]*Id="${rId}"[^/]*/>`), '');

    // Add the correct one
    xml = xml.replace('</Relationships>',
        `<Relationship Id="${rId}" Type="${type}" Target="${target}"/></Relationships>`);
    zip.file(relsPath, xml);
}

// Add new rel (auto-generate rId), return rId
async function addRelToFile(zip, relsPath, target, type) {
    let xml = zip.file(relsPath)
        ? await zip.file(relsPath).async('string')
        : relsDoc('');

    // Reuse if target already present
    const escaped = target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const existing = xml.match(new RegExp(`<Relationship[^>]*Target="${escaped}"[^/]*/>`));
    if (existing) {
        const m = existing[0].match(/Id="([^"]+)"/);
        if (m) return m[1];
    }

    const ids = [...xml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1]));
    const rId = `rId${ids.length ? Math.max(...ids) + 1 : 1}`;
    xml = xml.replace('</Relationships>',
        `<Relationship Id="${rId}" Type="${type}" Target="${target}"/></Relationships>`);
    zip.file(relsPath, xml);
    return rId;
}

// Add header/footer file rel to document.xml.rels
async function addRelToDocRels(zip, target, type, rId) {
    const path = 'word/_rels/document.xml.rels';
    let xml = zip.file(path) ? await zip.file(path).async('string') : relsDoc('');
    if (xml.includes(`Id="${rId}"`)) return;
    xml = xml.replace('</Relationships>',
        `<Relationship Id="${rId}" Type="${type}" Target="${target}"/></Relationships>`);
    zip.file(path, xml);
}

// Wire header/footerReference into document.xml sectPr
async function wireRef(zip, kind, rId) {
    let xml = await zip.file('word/document.xml').async('string');
    const tag = kind === 'header' ? 'w:headerReference' : 'w:footerReference';
    const ref = `<${tag} w:type="default" r:id="${rId}"/>`;
    if (xml.includes(ref)) return;
    if (xml.includes('</w:sectPr>')) {
        xml = xml.replace(/<\/w:sectPr>/, `${ref}</w:sectPr>`);
    } else {
        xml = xml.replace(/<\/w:body>/, `<w:sectPr>${ref}</w:sectPr></w:body>`);
    }
    zip.file('word/document.xml', xml);
}

// ─────────────────────────────────────────────────────
// Content type helpers
// ─────────────────────────────────────────────────────
async function ensureContentType(zip, ext, ct) {
    let xml = await zip.file('[Content_Types].xml').async('string');
    if (!xml.includes(`Extension="${ext}"`)) {
        xml = xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${ct}"/></Types>`);
        zip.file('[Content_Types].xml', xml);
    }
}

async function ensureOverride(zip, part, ct) {
    let xml = await zip.file('[Content_Types].xml').async('string');
    if (!xml.includes(`PartName="${part}"`)) {
        xml = xml.replace('</Types>', `<Override PartName="${part}" ContentType="${ct}"/></Types>`);
        zip.file('[Content_Types].xml', xml);
    }
}


