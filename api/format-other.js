// api/format-other.js — Vercel Serverless Function
// "Other Document" mode: preserves ALL original content exactly as-is.
// Only patches in header logo, footer logo, and/or page border via XML surgery.
// No content is extracted, parsed, or rebuilt — the raw XML is preserved.

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
        const outputBuffer = await applyOtherDocFormatting(inputBuffer, opts);
        const outputBase64 = outputBuffer.toString('base64');
        const baseName = (filename || 'document').replace(/\.docx$/i, '');

        return res.status(200).json({
            ok: true,
            docx: outputBase64,
            filename: `${baseName}_Formatted.docx`,
            // Preview just reports what was applied
            preview: {
                headerLogo: opts.headerLogo,
                footerLogo: opts.footerLogo,
                pageBorder: opts.pageBorder,
            }
        });

    } catch (err) {
        return res.status(500).json({ error: err.message });
    }
}

// ═══════════════════════════════════════════════════════
// Core: patch the docx ZIP without touching body content
// ═══════════════════════════════════════════════════════
async function applyOtherDocFormatting(buffer, opts) {
    const zip = await JSZip.loadAsync(buffer);
    const logoBuffer = Buffer.from(LOGO_B64, 'base64');

    // ── 1. Page Border ──
    if (opts.pageBorder) {
        await applyPageBorder(zip);
    }

    // ── 2. Header Logo ──
    if (opts.headerLogo) {
        await injectHeaderLogo(zip, logoBuffer);
    }

    // ── 3. Footer Logo ──
    if (opts.footerLogo) {
        await injectFooterLogo(zip, logoBuffer);
    }

    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ─────────────────────────────────────────────────────
// Apply triple dark-blue page border via sectPr XML
// ─────────────────────────────────────────────────────
async function applyPageBorder(zip) {
    let docXml = await zip.file('word/document.xml').async('string');

    const borderXml = `<w:pgBorders w:offsetFrom="page">` +
        `<w:top w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:left w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:bottom w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `<w:right w:val="triple" w:sz="24" w:space="1" w:color="1F3864"/>` +
        `</w:pgBorders>`;

    // Remove any existing pgBorders first
    docXml = docXml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/g, '');

    // Insert before </w:sectPr>
    if (docXml.includes('</w:sectPr>')) {
        docXml = docXml.replace(/<\/w:sectPr>/g, borderXml + '</w:sectPr>');
    } else {
        // No sectPr — inject one at the end of the body
        docXml = docXml.replace(/<\/w:body>/, `<w:sectPr>${borderXml}</w:sectPr></w:body>`);
    }

    zip.file('word/document.xml', docXml);
}

// ─────────────────────────────────────────────────────
// Inject XtremeLabs logo into document header
// Creates header1.xml + relationships if they don't exist
// ─────────────────────────────────────────────────────
async function injectHeaderLogo(zip, logoBuffer) {
    // Add the image to media/
    const imgName = 'xtremelabs_logo_h.jpg';
    zip.file(`word/media/${imgName}`, logoBuffer);

    // Ensure content type exists
    await ensureContentType(zip, 'jpg', 'image/jpeg');

    // Determine which header files exist
    const headerFiles = Object.keys(zip.files).filter(f => /^word\/header\d*\.xml$/.test(f));

    if (headerFiles.length === 0) {
        // No headers — create header1.xml and wire it up
        const rId = await addRelationship(zip, 'word/_rels/document.xml.rels', `media/${imgName}`, 'image');
        const headerRId = 'rId_hdr1';
        const headerXml = buildHeaderXml(rId);
        zip.file('word/header1.xml', headerXml);

        // Create header1.xml.rels
        zip.file('word/_rels/header1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);

        // Wire header ref in document.xml sectPr
        await wireHeaderRef(zip, headerRId, 'word/header1.xml');

        // Add document rel for the header file itself
        await addFileRelationship(zip, 'word/_rels/document.xml.rels', 'header1.xml',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', headerRId);

        // Add content type for header
        await ensureOverride(zip, '/word/header1.xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    } else {
        // Patch existing header(s)
        for (const hPath of headerFiles) {
            const relsPath = hPath.replace('word/', 'word/_rels/') + '.rels';
            const imgRId = await addRelToFile(zip, relsPath, `../media/${imgName}`, 'image');
            let hXml = await zip.file(hPath).async('string');
            hXml = injectLogoIntoHeaderXml(hXml, imgRId, 'header');
            zip.file(hPath, hXml);
        }
    }
}

// ─────────────────────────────────────────────────────
// Inject footer logo
// ─────────────────────────────────────────────────────
async function injectFooterLogo(zip, logoBuffer) {
    const imgName = 'xtremelabs_logo_f.jpg';
    zip.file(`word/media/${imgName}`, logoBuffer);
    await ensureContentType(zip, 'jpg', 'image/jpeg');

    const footerFiles = Object.keys(zip.files).filter(f => /^word\/footer\d*\.xml$/.test(f));

    if (footerFiles.length === 0) {
        const rId = await addRelationship(zip, 'word/_rels/document.xml.rels', `media/${imgName}`, 'image');
        const footerRId = 'rId_ftr1';
        const footerXml = buildFooterXml(rId);
        zip.file('word/footer1.xml', footerXml);
        zip.file('word/_rels/footer1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
        await wireFooterRef(zip, footerRId, 'word/footer1.xml');
        await addFileRelationship(zip, 'word/_rels/document.xml.rels', 'footer1.xml',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', footerRId);
        await ensureOverride(zip, '/word/footer1.xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
    } else {
        for (const fPath of footerFiles) {
            const relsPath = fPath.replace('word/', 'word/_rels/') + '.rels';
            const imgRId = await addRelToFile(zip, relsPath, `../media/${imgName}`, 'image');
            let fXml = await zip.file(fPath).async('string');
            fXml = injectLogoIntoHeaderXml(fXml, imgRId, 'footer');
            zip.file(fPath, fXml);
        }
    }
}

// ─────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────

function buildHeaderXml(imgRId) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"` +
    ` xmlns:v="urn:schemas-microsoft-com:vml"` +
    ` xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"` +
    ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
    ` xmlns:w10="urn:schemas-microsoft-com:office:word"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
    ` xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"` +
    ` xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"` +
    ` xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"` +
    ` xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">` +
    buildLogoParaXml(imgRId, 'header') +
    `</w:hdr>`;
}

function buildFooterXml(imgRId) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"` +
    ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
    ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"` +
    ` xmlns:v="urn:schemas-microsoft-com:vml"` +
    ` xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"` +
    ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
    ` xmlns:w10="urn:schemas-microsoft-com:office:word"` +
    ` xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
    ` xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"` +
    ` xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"` +
    ` xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"` +
    ` xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">` +
    buildLogoParaXml(imgRId, 'footer') +
    `</w:ftr>`;
}

// Build an inline image paragraph for header/footer
function buildLogoParaXml(rId, mode) {
    // Header: right-aligned logo (tab to right). Footer: centered with "Powered By:"
    const EMU_W = mode === 'header' ? 1270000 : 857250;  // ~140px / ~95px at 96dpi
    const EMU_H = mode === 'header' ?  292100 : 190500;  // ~32px  / ~21px

    const imgXml = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
        `<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU_W}" cy="${EMU_H}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic>`;

    const drawingXml = `<w:drawing>` +
        `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
        `<wp:extent cx="${EMU_W}" cy="${EMU_H}"/>` +
        `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        `<wp:docPr id="1" name="logo"/>` +
        `<wp:cNvGraphicFramePr/>` +
        imgXml +
        `</wp:inline></w:drawing>`;

    if (mode === 'header') {
        // Right-aligned via tab stop
        return `<w:p>` +
            `<w:pPr><w:jc w:val="right"/></w:pPr>` +
            `<w:r>${drawingXml}</w:r>` +
            `</w:p>`;
    } else {
        // Centered with label
        return `<w:p>` +
            `<w:pPr><w:jc w:val="center"/>` +
            `<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="2E5FA3"/></w:pBdr>` +
            `</w:pPr>` +
            `<w:r><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="20"/></w:rPr>` +
            `<w:t xml:space="preserve">Powered By:  </w:t></w:r>` +
            `<w:r>${drawingXml}</w:r>` +
            `</w:p>`;
    }
}

// Inject a logo paragraph into an existing header/footer XML (prepend for header, append for footer)
function injectLogoIntoHeaderXml(xml, rId, mode) {
    const logoPara = buildLogoParaXml(rId, mode);
    if (mode === 'header') {
        // Prepend inside root element
        return xml.replace(/(<w:hdr[^>]*>)/, `$1${logoPara}`);
    } else {
        // Append before closing root element
        return xml.replace(/<\/w:ftr>/, `${logoPara}</w:ftr>`);
    }
}

// Add a relationship to a .rels file (document.xml.rels level), returns rId
async function addRelationship(zip, relsPath, target, type) {
    const typeMap = {
        image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
        footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
    };
    const fullType = typeMap[type] || type;
    let relsXml = '';
    if (zip.file(relsPath)) {
        relsXml = await zip.file(relsPath).async('string');
    } else {
        relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }

    // Check if relationship to this target already exists
    const existingMatch = relsXml.match(new RegExp(`Target="${target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^/]*/>`));
    if (existingMatch) {
        const idMatch = existingMatch[0].match(/Id="([^"]+)"/);
        if (idMatch) return idMatch[1];
    }

    // Find next available rId number
    const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1]));
    const nextId = ids.length ? Math.max(...ids) + 1 : 100;
    const rId = `rId${nextId}`;

    const newRel = `<Relationship Id="${rId}" Type="${fullType}" Target="${target}"/>`;
    relsXml = relsXml.replace('</Relationships>', `${newRel}</Relationships>`);
    zip.file(relsPath, relsXml);
    return rId;
}

// Add rel to a header/footer rels file (relative target)
async function addRelToFile(zip, relsPath, target, type) {
    const typeMap = {
        image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    };
    const fullType = typeMap[type] || type;

    let relsXml = '';
    if (zip.file(relsPath)) {
        relsXml = await zip.file(relsPath).async('string');
    } else {
        relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }

    const escapedTarget = target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const existingMatch = relsXml.match(new RegExp(`Target="${escapedTarget}"[^/]*/>`));
    if (existingMatch) {
        const idMatch = existingMatch[0].match(/Id="([^"]+)"/);
        if (idMatch) return idMatch[1];
    }

    const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1]));
    const nextId = ids.length ? Math.max(...ids) + 1 : 10;
    const rId = `rId${nextId}`;

    const newRel = `<Relationship Id="${rId}" Type="${fullType}" Target="${target}"/>`;
    relsXml = relsXml.replace('</Relationships>', `${newRel}</Relationships>`);
    zip.file(relsPath, relsXml);
    return rId;
}

// Add a file-level relationship with a specific rId
async function addFileRelationship(zip, relsPath, target, type, rId) {
    let relsXml = zip.file(relsPath) ? await zip.file(relsPath).async('string') : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    if (relsXml.includes(`Id="${rId}"`)) return; // already there
    const newRel = `<Relationship Id="${rId}" Type="${type}" Target="${target}"/>`;
    relsXml = relsXml.replace('</Relationships>', `${newRel}</Relationships>`);
    zip.file(relsPath, relsXml);
}

// Wire a headerReference into sectPr of document.xml
async function wireHeaderRef(zip, rId, filePath) {
    let docXml = await zip.file('word/document.xml').async('string');
    const ref = `<w:headerReference w:type="default" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
    if (!docXml.includes(ref)) {
        docXml = docXml.replace(/<\/w:sectPr>/, `${ref}</w:sectPr>`);
        if (!docXml.includes('</w:sectPr>')) {
            docXml = docXml.replace(/<\/w:body>/, `<w:sectPr>${ref}</w:sectPr></w:body>`);
        }
    }
    zip.file('word/document.xml', docXml);
}

// Wire a footerReference into sectPr of document.xml
async function wireFooterRef(zip, rId, filePath) {
    let docXml = await zip.file('word/document.xml').async('string');
    const ref = `<w:footerReference w:type="default" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`;
    if (!docXml.includes(ref)) {
        if (docXml.includes('</w:sectPr>')) {
            docXml = docXml.replace(/<\/w:sectPr>/, `${ref}</w:sectPr>`);
        } else {
            docXml = docXml.replace(/<\/w:body>/, `<w:sectPr>${ref}</w:sectPr></w:body>`);
        }
    }
    zip.file('word/document.xml', docXml);
}

// Ensure an image content type entry exists in [Content_Types].xml
async function ensureContentType(zip, ext, contentType) {
    let ctXml = await zip.file('[Content_Types].xml').async('string');
    if (!ctXml.includes(`Extension="${ext}"`)) {
        ctXml = ctXml.replace('</Types>', `<Default Extension="${ext}" ContentType="${contentType}"/></Types>`);
        zip.file('[Content_Types].xml', ctXml);
    }
}

// Ensure an Override entry exists in [Content_Types].xml
async function ensureOverride(zip, partName, contentType) {
    let ctXml = await zip.file('[Content_Types].xml').async('string');
    if (!ctXml.includes(`PartName="${partName}"`)) {
        ctXml = ctXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`);
        zip.file('[Content_Types].xml', ctXml);
    }
}
