// api/format-other.js — Vercel Serverless Function
// "Other Document" mode: preserves ALL original content exactly as-is.
// Only patches in header logo, footer logo, and/or page border via XML surgery.
// No content is extracted, parsed, or rebuilt — the raw XML is preserved.

import JSZip from 'jszip';

const LOGO_B64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAA+AXoDASIAAhEBAxEB/8QAHQAAAgMBAQEBAQAAAAAAAAAAAAcGCAkFBAMCAf/EAFEQAAEDAwICBAYMCwYDCQAAAAECAwQABQYHERIhCDFBURMYImGU0gkUFlNWcXWBkZOz0RUjMzdCVVeSlbHTMjZScoKhQ1RifWVvEW4TEU1WYqWxMZPpxuK3MYHi7+PyMWf/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAA+AXoDASIAAhEBAxEB/8QAHQAAAgMBAQEBAQAAAAAAAAAAAAcGCAkFBAMCAf/EAFEQAAEDAwICBAYMCwYDCQAAAAECAwQABQYHERIhCDFBURMYImGU0gkUFlNWcXWBkZOz0RUjMzdCVVeSlbHTMjZScoKhQ1RifWVvEW4TEU1WYqWxMZPpxuK3MYHi7+PyMWf/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAA+AXoDASIAAhEBAxEB/8QAHQAAAgMBAQEBAQAAAAAAAAAAAAcGCAkFBAMCAf/EAFEQAAEDAwICBAYMCwYDCQAAAAECAwQABQYHERIhCDFBURMYImGU0gkUFlNWcXWBkZOz0RUjMzdCVVeSlbHTMjZScoKhQ1RifWVvEW4TEU1WYqWxMZPpxuK3MYHi7+PyMWc=';

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
