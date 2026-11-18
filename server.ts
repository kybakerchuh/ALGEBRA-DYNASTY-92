import 'dotenv/config';
import express, { type Request, type Response } from 'express';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, Type, type Part } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);

function getGeminiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not configured on the server.');
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

function getBearerToken(req: Request): string | null {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) return null;
  const token = authHeader.slice('Bearer '.length).trim();
  return token.length > 10 ? token : null;
}

function extractSpreadsheetId(input: string): string | null {
  const trimmed = (input || '').trim();
  if (!trimmed) return null;
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  if (/^[a-zA-Z0-9_-]{15,120}$/.test(trimmed)) return trimmed;
  return null;
}

interface ExtractedLink {
  url: string;
  label: string;
  kind: 'doc' | 'slides' | 'sheet' | 'drive' | 'web';
  fileId?: string;
}

function classifyUrl(rawUrl: string, label = ''): ExtractedLink {
  const url = rawUrl.trim();
  const docMatch = url.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (docMatch?.[1]) {
    return { url, label: label || 'Google Doc', kind: 'doc', fileId: docMatch[1] };
  }
  const slidesMatch = url.match(/docs\.google\.com\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (slidesMatch?.[1]) {
    return { url, label: label || 'Google Slides', kind: 'slides', fileId: slidesMatch[1] };
  }
  const sheetMatch = url.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (sheetMatch?.[1]) {
    return { url, label: label || 'Google Sheet', kind: 'sheet', fileId: sheetMatch[1] };
  }
  const driveFileMatch =
    url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/) ||
    url.match(/drive\.google\.com\/(?:open|uc)\?[^#]*\bid=([a-zA-Z0-9_-]+)/);
  if (driveFileMatch?.[1]) {
    return { url, label: label || 'Drive File', kind: 'drive', fileId: driveFileMatch[1] };
  }
  return { url, label: label || url, kind: 'web' };
}

interface GridCellData {
  formattedValue?: string;
  hyperlink?: string;
  textFormatRuns?: Array<{
    format?: {
      link?: {
        uri?: string;
      };
    };
  }>;
}

function extractCellTextAndLinks(cell: GridCellData | undefined): {
  text: string;
  links: ExtractedLink[];
} {
  if (!cell) return { text: '', links: [] };
  const text = (cell.formattedValue || '').trim();
  const seenUrls = new Set<string>();
  const links: ExtractedLink[] = [];

  const pushUrl = (raw: string | undefined, lbl: string) => {
    if (!raw) return;
    const clean = raw.trim();
    if (!/^https?:\/\//i.test(clean) || seenUrls.has(clean)) return;
    seenUrls.add(clean);
    links.push(classifyUrl(clean, lbl));
  };

  if (cell.hyperlink) {
    pushUrl(cell.hyperlink, text || 'Link');
  }
  if (Array.isArray(cell.textFormatRuns)) {
    for (const run of cell.textFormatRuns) {
      if (run?.format?.link?.uri) {
        pushUrl(run.format.link.uri, text || 'Link');
      }
    }
  }
  const urlRegex = /https?:\/\/[^\s<>"')\]}]+/gi;
  const matches = text.match(urlRegex);
  if (matches) {
    for (const m of matches) {
      pushUrl(m, text === m ? '' : text.replace(m, '').trim());
    }
  }

  return { text, links };
}

function extractGoogleDocText(docJson: any): string {
  const parts: string[] = [];
  if (docJson?.title) parts.push(`Title: ${docJson.title}`);

  const readStructuralElements = (elements: any[]) => {
    if (!Array.isArray(elements)) return;
    for (const el of elements) {
      if (el.paragraph?.elements) {
        const line = el.paragraph.elements
          .map((pe: any) => pe.textRun?.content || '')
          .join('')
          .trim();
        if (line) parts.push(line);
      } else if (el.table?.tableRows) {
        for (const row of el.table.tableRows) {
          const cells: string[] = [];
          for (const cell of row.tableCells || []) {
            const cellParts: string[] = [];
            for (const cEl of cell.content || []) {
              if (cEl.paragraph?.elements) {
                const t = cEl.paragraph.elements
                  .map((pe: any) => pe.textRun?.content || '')
                  .join('')
                  .trim();
                if (t) cellParts.push(t);
              }
            }
            if (cellParts.length) cells.push(cellParts.join(' '));
          }
          if (cells.length) parts.push(cells.join(' | '));
        }
      }
    }
  };

  readStructuralElements(docJson?.body?.content || []);
  return parts.join('\n').slice(0, 18000);
}

function extractGoogleSlidesText(presJson: any): string {
  const parts: string[] = [];
  if (presJson?.title) parts.push(`Presentation: ${presJson.title}`);

  const extractShapeTexts = (pageElements: any[]): string[] => {
    const out: string[] = [];
    if (!Array.isArray(pageElements)) return out;
    for (const pe of pageElements) {
      if (pe.shape?.text?.textElements) {
        const txt = pe.shape.text.textElements
          .map((te: any) => te.textRun?.content || '')
          .join('')
          .trim();
        if (txt) out.push(txt);
      }
      if (pe.table?.tableRows) {
        for (const row of pe.table.tableRows) {
          const rowText = (row.tableCells || [])
            .map((cell: any) =>
              (cell.text?.textElements || [])
                .map((te: any) => te.textRun?.content || '')
                .join('')
                .trim()
            )
            .filter(Boolean)
            .join(' | ');
          if (rowText) out.push(rowText);
        }
      }
      if (pe.elementGroup?.children) {
        out.push(...extractShapeTexts(pe.elementGroup.children));
      }
    }
    return out;
  };

  const slides = Array.isArray(presJson?.slides) ? presJson.slides : [];
  slides.forEach((slide: any, idx: number) => {
    const slideTexts = extractShapeTexts(slide.pageElements || []);
    const notesTexts = extractShapeTexts(slide.slideProperties?.notesPage?.pageElements || []);
    if (slideTexts.length || notesTexts.length) {
      parts.push(`--- Slide ${idx + 1} ---`);
      if (slideTexts.length) parts.push(slideTexts.join('\n'));
      if (notesTexts.length) parts.push(`Speaker Notes: ${notesTexts.join(' ')}`);
    }
  });

  return parts.join('\n').slice(0, 18000);
}

function stripHtmlToText(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

async function fetchGoogleApiJson(url: string, accessToken: string): Promise<{ ok: boolean; status: number; data: any; errorText?: string }> {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const errorText = await res.text().catch(() => '');
    return { ok: false, status: res.status, data: null, errorText };
  }
  const data = await res.json();
  return { ok: true, status: res.status, data };
}

function slugifyTopicId(raw: string, fallbackIdx: number): string {
  const base = (raw || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 72);
  return base.length >= 2 ? `topic_${base}` : `topic_alg_${fallbackIdx}`;
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '6mb' }));

  app.get('/api/health', (_req: Request, res: Response) => {
    res.json({ ok: true });
  });

  // 1. List recent Google Sheets from the teacher's Drive
  app.post('/api/workspace/list-sheets', async (req: Request, res: Response) => {
    const token = getBearerToken(req);
    if (!token) {
      res.status(401).json({ error: 'Missing Google Workspace access token. Please connect Google Sheets.' });
      return;
    }

    try {
      const q = encodeURIComponent("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
      const url = `https://www.googleapis.com/drive/v3/files?q=${q}&orderBy=modifiedTime%20desc&pageSize=25&supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,modifiedTime,webViewLink)`;
      const resp = await fetchGoogleApiJson(url, token);
      if (!resp.ok) {
        res.status(resp.status).json({ error: `Google Drive error (${resp.status}): ${resp.errorText?.slice(0, 200)}` });
        return;
      }
      res.json({ files: resp.data.files || [] });
    } catch (err: any) {
      res.status(500).json({ error: err?.message || 'Failed to list Google Sheets.' });
    }
  });

  // 2. Inspect a Google Sheet (metadata, tab list, headers, rows + hyperlinks)
  app.post('/api/workspace/inspect-sheet', async (req: Request, res: Response) => {
    const token = getBearerToken(req);
    if (!token) {
      res.status(401).json({ error: 'Missing Google Workspace access token. Please sign in with Google Teacher Sync.' });
      return;
    }

    const { sheetUrlOrId, tabTitle, headerRowIndex } = req.body || {};
    const spreadsheetId = extractSpreadsheetId(String(sheetUrlOrId || ''));
    if (!spreadsheetId) {
      res.status(400).json({ error: 'Please paste a valid Google Sheets URL or spreadsheet ID.' });
      return;
    }

    try {
      const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets(properties(sheetId,title,gridProperties))`;
      const metaResp = await fetchGoogleApiJson(metaUrl, token);
      if (!metaResp.ok) {
        res.status(metaResp.status).json({
          error:
            metaResp.status === 403 || metaResp.status === 401
              ? 'Permission denied reading that Google Sheet. Make sure you signed in with the account that has access to it.'
              : `Could not open spreadsheet (${metaResp.status}).`,
        });
        return;
      }

      const title = metaResp.data?.properties?.title || 'Untitled Spreadsheet';
      const tabs: Array<{ sheetId: number; title: string }> = (metaResp.data?.sheets || []).map((s: any) => ({
        sheetId: s.properties?.sheetId ?? 0,
        title: s.properties?.title || 'Sheet1',
      }));

      if (!tabs.length) {
        res.status(400).json({ error: 'This spreadsheet has no tabs.' });
        return;
      }

      const activeTab =
        (tabTitle && tabs.find((t) => t.title === tabTitle)?.title) || tabs[0].title;

      const safeRange = `'${activeTab.replace(/'/g, "''")}'!A1:Z400`;
      const gridUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?ranges=${encodeURIComponent(
        safeRange
      )}&includeGridData=true&fields=sheets(data(rowData(values(formattedValue,hyperlink,textFormatRuns))))`;
      const gridResp = await fetchGoogleApiJson(gridUrl, token);
      if (!gridResp.ok) {
        res.status(gridResp.status).json({ error: `Could not read sheet tab "${activeTab}" (${gridResp.status}).` });
        return;
      }

      const rawRows: any[] = gridResp.data?.sheets?.[0]?.data?.[0]?.rowData || [];
      const parsedGrid: Array<Array<{ text: string; links: ExtractedLink[] }>> = rawRows.map((row) =>
        (row.values || []).map((cell: GridCellData) => extractCellTextAndLinks(cell))
      );

      const maxCols = parsedGrid.reduce((m, r) => Math.max(m, r.length), 0);
      if (maxCols === 0) {
        res.status(400).json({ error: `Tab "${activeTab}" appears to be empty.` });
        return;
      }

      let detectedHeaderIdx = Number.isInteger(headerRowIndex) ? Math.max(0, headerRowIndex) : -1;
      if (detectedHeaderIdx < 0) {
        detectedHeaderIdx = parsedGrid.findIndex((row) => row.filter((c) => c.text.length > 0).length >= 2);
        if (detectedHeaderIdx < 0) detectedHeaderIdx = 0;
      }

      const headerRow = parsedGrid[detectedHeaderIdx] || [];
      const headers = Array.from({ length: maxCols }, (_, colIdx) => {
        const raw = headerRow[colIdx]?.text || '';
        return raw || `Column ${String.fromCharCode(65 + (colIdx % 26))}`;
      });

      const dataRows: Array<{
        rowIndex: number;
        cells: Array<{ text: string; links: ExtractedLink[] }>;
      }> = [];

      for (let i = detectedHeaderIdx + 1; i < parsedGrid.length; i++) {
        const row = parsedGrid[i] || [];
        const hasAnyContent = row.some((c) => c.text.length > 0 || c.links.length > 0);
        if (!hasAnyContent) continue;
        const padded = Array.from({ length: maxCols }, (_, colIdx) => row[colIdx] || { text: '', links: [] });
        dataRows.push({ rowIndex: i + 1, cells: padded });
      }

      // Auto-suggest columns for Standards, Topic hints, and Resource links
      const standardCol = headers.findIndex((h) => /standard|code|benchmark|ccss|teks|sol|learning target/i.test(h));
      const topicHintCol = headers.findIndex((h) => /topic|unit|lesson|title|skill|concept|chapter|module/i.test(h));

      const resourceCols: number[] = [];
      for (let c = 0; c < maxCols; c++) {
        const headerLooksResource = /link|url|doc|slide|pdf|resource|notes|practice|activity|worksheet|online|video|material/i.test(headers[c]);
        const hasLinksInCol = dataRows.some((r) => r.cells[c]?.links?.length > 0);
        if (hasLinksInCol || (headerLooksResource && c !== standardCol)) {
          resourceCols.push(c);
        }
      }

      res.json({
        spreadsheetId,
        title,
        tabs,
        activeTab,
        headerRowIndex: detectedHeaderIdx,
        headers,
        rows: dataRows.slice(0, 250),
        suggestedMapping: {
          standardCol: standardCol >= 0 ? standardCol : 0,
          topicHintCol: topicHintCol >= 0 ? topicHintCol : -1,
          resourceCols,
        },
      });
    } catch (err: any) {
      res.status(500).json({ error: err?.message || 'Unexpected error inspecting Google Sheet.' });
    }
  });

  // 3. One-pass whole-sheet clustering: Standards + Resources -> Deduplicated Student-Friendly Units & Topics
  app.post('/api/workspace/cluster-topics', async (req: Request, res: Response) => {
    try {
      const { rows, headers, standardCol, topicHintCol, resourceCols, sheetTitle } = req.body || {};
      if (!Array.isArray(rows) || rows.length === 0) {
        res.status(400).json({ error: 'No rows provided to organize.' });
        return;
      }

      const compactRows = rows.slice(0, 180).map((r: any, idx: number) => {
        const stdText = standardCol >= 0 ? (r.cells?.[standardCol]?.text || '') : '';
        const hintText = topicHintCol >= 0 ? (r.cells?.[topicHintCol]?.text || '') : '';
        const extraTexts: string[] = [];
        const links: ExtractedLink[] = [];

        const colsToScan: number[] =
          Array.isArray(resourceCols) && resourceCols.length > 0
            ? resourceCols
            : (r.cells || []).map((_: any, cIdx: number) => cIdx);

        for (const cIdx of colsToScan) {
          const cell = r.cells?.[cIdx];
          if (!cell) continue;
          if (Array.isArray(cell.links)) {
            for (const l of cell.links) {
              links.push({
                url: l.url,
                label: l.label || headers?.[cIdx] || l.kind,
                kind: l.kind,
                fileId: l.fileId,
              });
            }
          }
          if (cell.text && cIdx !== standardCol && cIdx !== topicHintCol) {
            extraTexts.push(`${headers?.[cIdx] || 'Info'}: ${cell.text.slice(0, 120)}`);
          }
        }

        // Also scan all remaining cells for any unmapped hyperlinks so nothing is missed
        if (Array.isArray(r.cells)) {
          r.cells.forEach((cell: any, cIdx: number) => {
            if (!colsToScan.includes(cIdx) && Array.isArray(cell?.links)) {
              for (const l of cell.links) {
                if (!links.some((existing) => existing.url === l.url)) {
                  links.push(l);
                }
              }
            }
          });
        }

        return {
          idx,
          rowNumber: r.rowIndex || idx + 2,
          standard: stdText.slice(0, 260),
          topicHint: hintText.slice(0, 140),
          extra: extraTexts.join(' | ').slice(0, 220),
          resourceCount: links.length,
          resourceLabels: links.map((l) => `${l.kind}:${l.label}`.slice(0, 60)),
          links,
        };
      });

      const ai = getGeminiClient();
      const promptInput = compactRows.map((r) => ({
        idx: r.idx,
        standard: r.standard,
        topicHint: r.topicHint,
        extra: r.extra,
        resourceLabels: r.resourceLabels,
      }));

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: `You are an expert Algebra curriculum designer helping a high school teacher organize a spreadsheet ("${sheetTitle || 'Algebra Resources'}") into a clean, student-friendly review topic menu.

Here are all the rows from the spreadsheet (each row has a curriculum standard and/or resource info):
${JSON.stringify(promptInput)}

TASK:
1. Group ALL rows together in ONE unified curriculum map organized into these exact Algebra OST Playbook Units in order (only include units matching the rows present):
   - UnitOrder 1: "EXPRESSIONS, EQUATIONS, AND INEQUALITIES"
   - UnitOrder 2: "FUNCTIONS AND FUNCTION NOTATIONS"
   - UnitOrder 3: "CREATING, GRAPHING, MODELING LINEAR"
   - UnitOrder 4: "CREATING, SOLVING, MODELING SYSTEMS OF EQUATIONS"
   - UnitOrder 5: "ALL THINGS POLYNOMIALS"
   - UnitOrder 6: "FACTORING, SOLVING, MODELING, ANALYZING QUADRATICS"
   - UnitOrder 7: "CREATE, ANALYZE, MODEL EXPONENTIAL FUNCTIONS"
   - UnitOrder 8: "DATA ANALYSIS (BOX WHISKER, LINE OF BEST FIT, DATA, ETC...)"
2. Translate technical standard codes/descriptions into clear, student-friendly Topic Names (e.g., "Solving Multi-Step Equations", "Graphing Slope-Intercept Form", "Factoring Quadratics"). Do NOT put cryptic standard codes (like HSA.REI.B.4) in topicTitle—put the standards in the standards array!
3. MERGE closely related or overlapping rows/standards into ONE unified topic so there are ZERO duplicates and closely related skills do NOT end up scattered across different spots.
4. Keep topics focused and digestible (typically 1-4 related rows per topic; aim for roughly 6 to 18 total topics depending on the sheet size).
5. Every non-empty row index (idx from 0 to ${compactRows.length - 1}) MUST be assigned to at least one topic and never duplicated across different topics unless necessary.`,
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              units: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    unitOrder: { type: Type.INTEGER },
                    unitName: { type: Type.STRING },
                    topics: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.OBJECT,
                        properties: {
                          topicTitle: { type: Type.STRING },
                          studentSummary: { type: Type.STRING },
                          standards: {
                            type: Type.ARRAY,
                            items: { type: Type.STRING },
                          },
                          rowIndices: {
                            type: Type.ARRAY,
                            items: { type: Type.INTEGER },
                          },
                        },
                        required: ['topicTitle', 'studentSummary', 'standards', 'rowIndices'],
                      },
                    },
                  },
                  required: ['unitOrder', 'unitName', 'topics'],
                },
              },
            },
            required: ['units'],
          },
        },
      });

      const parsed = JSON.parse((response.text || '{"units":[]}').trim());
      const rawUnits = Array.isArray(parsed.units) ? parsed.units : [];
      const seenTopicIds = new Set<string>();
      const usedRowIdx = new Set<number>();

      const normalizedTopics: Array<{
        topicId: string;
        unitOrder: number;
        unitName: string;
        topicTitle: string;
        studentSummary: string;
        standards: string[];
        rowNumbers: number[];
        resources: ExtractedLink[];
      }> = [];

      let topicCounter = 1;
      for (const unit of rawUnits.sort((a: any, b: any) => (a.unitOrder || 99) - (b.unitOrder || 99))) {
        const unitName = String(unit.unitName || 'Algebra Unit').trim().slice(0, 120);
        const unitOrder = Math.max(1, Math.min(99, Math.trunc(Number(unit.unitOrder) || 1)));
        for (const t of unit.topics || []) {
          const topicTitle = String(t.topicTitle || `Algebra Topic ${topicCounter}`).trim().slice(0, 160);
          let topicId = slugifyTopicId(topicTitle, topicCounter);
          while (seenTopicIds.has(topicId)) {
            topicCounter += 1;
            topicId = slugifyTopicId(`${topicTitle}_${topicCounter}`, topicCounter);
          }
          seenTopicIds.add(topicId);
          topicCounter += 1;

          const rowIndices: number[] = Array.isArray(t.rowIndices)
            ? t.rowIndices.filter((idx: any) => Number.isInteger(idx) && idx >= 0 && idx < compactRows.length)
            : [];

          const resources: ExtractedLink[] = [];
          const seenUrl = new Set<string>();
          const rowNumbers: number[] = [];
          const collectedStds: string[] = Array.isArray(t.standards)
            ? t.standards.map((s: any) => String(s).trim()).filter(Boolean)
            : [];

          for (const rIdx of rowIndices) {
            usedRowIdx.add(rIdx);
            const rowObj = compactRows[rIdx];
            rowNumbers.push(rowObj.rowNumber);
            if (rowObj.standard && !collectedStds.includes(rowObj.standard.slice(0, 120))) {
              collectedStds.push(rowObj.standard.slice(0, 120));
            }
            for (const link of rowObj.links) {
              if (!seenUrl.has(link.url)) {
                seenUrl.add(link.url);
                resources.push(link);
              }
            }
          }

          const finalStds = (collectedStds.length ? collectedStds : [topicTitle])
            .map((s) => s.slice(0, 120))
            .slice(0, 25);

          normalizedTopics.push({
            topicId,
            unitOrder,
            unitName,
            topicTitle,
            studentSummary: String(t.studentSummary || `Practice ${topicTitle}`).trim().slice(0, 300),
            standards: finalStds,
            rowNumbers,
            resources,
          });
        }
      }

      res.json({
        topics: normalizedTopics,
        totalRowsMatched: usedRowIdx.size,
        totalRowsInput: compactRows.length,
      });
    } catch (err: any) {
      res.status(500).json({ error: err?.message || 'Failed to cluster spreadsheet topics.' });
    }
  });

  // 4. Fetch linked Docs/Slides/Drive PDFs/Public URLs & generate deduplicated Algebra questions
  app.post('/api/workspace/generate-topic-questions', async (req: Request, res: Response) => {
    const token = getBearerToken(req);
    const {
      topicTitle,
      unitName,
      studentSummary,
      standards,
      resources,
      existingPrompts,
      count,
    } = req.body || {};

    if (!topicTitle) {
      res.status(400).json({ error: 'Topic title is required.' });
      return;
    }

    const targetCount = Math.max(4, Math.min(18, Math.trunc(Number(count) || 10)));
    const resourceList: ExtractedLink[] = Array.isArray(resources) ? resources.slice(0, 6) : [];
    const textSnippets: string[] = [];
    const pdfParts: Part[] = [];
    const resourceReport: Array<{ label: string; kind: string; status: string; detail: string }> = [];

    for (const r of resourceList) {
      const linkInfo = r.kind ? r : classifyUrl(r.url || '', r.label || '');
      const label = linkInfo.label || linkInfo.url;

      try {
        if ((linkInfo.kind === 'doc' || linkInfo.kind === 'slides' || linkInfo.kind === 'sheet' || linkInfo.kind === 'drive') && !token) {
          resourceReport.push({
            label,
            kind: linkInfo.kind,
            status: 'skipped',
            detail: 'Google sign-in token required to read Drive/Docs/Slides.',
          });
          continue;
        }

        if (linkInfo.kind === 'doc' && linkInfo.fileId && token) {
          const docRes = await fetchGoogleApiJson(
            `https://docs.googleapis.com/v1/documents/${linkInfo.fileId}`,
            token
          );
          if (docRes.ok) {
            const text = extractGoogleDocText(docRes.data);
            if (text.length > 20) {
              textSnippets.push(`=== GOOGLE DOC (${label}) ===\n${text}`);
              resourceReport.push({ label, kind: 'doc', status: 'ok', detail: 'Read Google Doc content' });
            } else {
              resourceReport.push({ label, kind: 'doc', status: 'skipped', detail: 'Google Doc had no readable text' });
            }
          } else {
            resourceReport.push({ label, kind: 'doc', status: 'skipped', detail: `Google Doc access error (${docRes.status})` });
          }
        } else if (linkInfo.kind === 'slides' && linkInfo.fileId && token) {
          const slidesRes = await fetchGoogleApiJson(
            `https://slides.googleapis.com/v1/presentations/${linkInfo.fileId}`,
            token
          );
          if (slidesRes.ok) {
            const text = extractGoogleSlidesText(slidesRes.data);
            if (text.length > 20) {
              textSnippets.push(`=== GOOGLE SLIDES (${label}) ===\n${text}`);
              resourceReport.push({ label, kind: 'slides', status: 'ok', detail: 'Read Google Slides deck' });
            } else {
              resourceReport.push({ label, kind: 'slides', status: 'skipped', detail: 'Slides deck had no readable text' });
            }
          } else {
            resourceReport.push({ label, kind: 'slides', status: 'skipped', detail: `Google Slides access error (${slidesRes.status})` });
          }
        } else if (linkInfo.kind === 'drive' && linkInfo.fileId && token) {
          const metaRes = await fetchGoogleApiJson(
            `https://www.googleapis.com/drive/v3/files/${linkInfo.fileId}?supportsAllDrives=true&fields=id,name,mimeType,size`,
            token
          );
          if (!metaRes.ok) {
            resourceReport.push({ label, kind: 'drive', status: 'skipped', detail: `Drive file access error (${metaRes.status})` });
            continue;
          }
          const mime = String(metaRes.data?.mimeType || '');
          const fileName = String(metaRes.data?.name || label);
          const sizeBytes = Number(metaRes.data?.size || 0);

          if (mime === 'application/vnd.google-apps.document') {
            const docRes = await fetchGoogleApiJson(
              `https://docs.googleapis.com/v1/documents/${linkInfo.fileId}`,
              token
            );
            if (docRes.ok) {
              textSnippets.push(`=== GOOGLE DOC (${fileName}) ===\n${extractGoogleDocText(docRes.data)}`);
              resourceReport.push({ label: fileName, kind: 'doc', status: 'ok', detail: 'Read Google Doc via Drive' });
            }
          } else if (mime === 'application/vnd.google-apps.presentation') {
            const slidesRes = await fetchGoogleApiJson(
              `https://slides.googleapis.com/v1/presentations/${linkInfo.fileId}`,
              token
            );
            if (slidesRes.ok) {
              textSnippets.push(`=== GOOGLE SLIDES (${fileName}) ===\n${extractGoogleSlidesText(slidesRes.data)}`);
              resourceReport.push({ label: fileName, kind: 'slides', status: 'ok', detail: 'Read Google Slides via Drive' });
            }
          } else if (mime === 'application/pdf') {
            if (sizeBytes > 12 * 1024 * 1024) {
              resourceReport.push({ label: fileName, kind: 'pdf', status: 'skipped', detail: 'PDF larger than 12MB limit' });
            } else if (pdfParts.length < 2) {
              const binRes = await fetch(
                `https://www.googleapis.com/drive/v3/files/${linkInfo.fileId}?alt=media&supportsAllDrives=true`,
                { headers: { Authorization: `Bearer ${token}` } }
              );
              if (binRes.ok) {
                const buf = Buffer.from(await binRes.arrayBuffer());
                pdfParts.push({
                  inlineData: {
                    mimeType: 'application/pdf',
                    data: buf.toString('base64'),
                  },
                });
                resourceReport.push({ label: fileName, kind: 'pdf', status: 'ok', detail: 'Loaded PDF from Drive' });
              } else {
                resourceReport.push({ label: fileName, kind: 'pdf', status: 'skipped', detail: `PDF download error (${binRes.status})` });
              }
            }
          } else {
            resourceReport.push({ label: fileName, kind: 'drive', status: 'skipped', detail: `Unsupported Drive type (${mime})` });
          }
        } else if (linkInfo.kind === 'web' && linkInfo.url) {
          const controller = new AbortController();
          const timer = setTimeout(() => controller.abort(), 5500);
          try {
            const webRes = await fetch(linkInfo.url, {
              signal: controller.signal,
              headers: { 'User-Agent': 'Mozilla/5.0 (compatible; AlgebraDynasty/1.0)' },
            });
            clearTimeout(timer);
            const cType = (webRes.headers.get('content-type') || '').toLowerCase();
            if (!webRes.ok) {
              resourceReport.push({
                label,
                kind: 'web',
                status: 'skipped',
                detail: `Online link blocked/login-walled (HTTP ${webRes.status}) — used standard & other docs`,
              });
            } else if (cType.includes('application/pdf') && pdfParts.length < 2) {
              const buf = Buffer.from(await webRes.arrayBuffer());
              if (buf.byteLength <= 10 * 1024 * 1024) {
                pdfParts.push({
                  inlineData: {
                    mimeType: 'application/pdf',
                    data: buf.toString('base64'),
                  },
                });
                resourceReport.push({ label, kind: 'pdf', status: 'ok', detail: 'Loaded public PDF' });
              } else {
                resourceReport.push({ label, kind: 'pdf', status: 'skipped', detail: 'Public PDF exceeded 10MB' });
              }
            } else if (cType.includes('text/html') || cType.includes('text/plain')) {
              const html = await webRes.text();
              const clean = stripHtmlToText(html).slice(0, 10000);
              if (clean.length < 80 || /sign in|log in|enable javascript|access denied/i.test(clean.slice(0, 400))) {
                resourceReport.push({
                  label,
                  kind: 'web',
                  status: 'skipped',
                  detail: 'Login-walled or JS-only site — used standard & other docs',
                });
              } else {
                textSnippets.push(`=== ONLINE RESOURCE (${label}) ===\n${clean}`);
                resourceReport.push({ label, kind: 'web', status: 'ok', detail: 'Extracted text from public page' });
              }
            } else {
              resourceReport.push({ label, kind: 'web', status: 'skipped', detail: 'Interactive/non-text link — used standard' });
            }
          } catch {
            clearTimeout(timer);
            resourceReport.push({
              label,
              kind: 'web',
              status: 'skipped',
              detail: 'Could not reach external site — used standard & other docs',
            });
          }
        }
      } catch (resErr: any) {
        resourceReport.push({
          label,
          kind: linkInfo.kind || 'unknown',
          status: 'skipped',
          detail: resErr?.message || 'Error reading resource',
        });
      }
    }

    try {
      const ai = getGeminiClient();
      const stdList = Array.isArray(standards) && standards.length ? standards : [topicTitle];
      const avoidList = Array.isArray(existingPrompts) ? existingPrompts.slice(0, 40) : [];

      const textPrompt = `You are a master High School Algebra teacher writing crisp, accurate, engaging review questions for a football-themed classroom review game ("Algebra Dynasty '92").

UNIT: ${unitName || 'Algebra'}
STUDENT TOPIC: ${topicTitle}
TOPIC SUMMARY: ${studentSummary || ''}
CURRICULUM STANDARDS TO COVER:
${stdList.map((s: string) => `- ${s}`).join('\n')}

${textSnippets.length > 0 ? `TEACHER RESOURCE MATERIALS:\n${textSnippets.join('\n\n').slice(0, 26000)}` : 'Note: Rely closely on the curriculum standards above (and any attached PDFs) to craft authentic Algebra problems.'}

${avoidList.length > 0 ? `DO NOT REPEAT OR CLOSELY PARAPHRASE THESE EXISTING QUESTIONS:\n${avoidList.map((p: string) => `- ${p}`).join('\n')}` : ''}

REQUIREMENTS:
1. Generate exactly ${targetCount} distinct, non-duplicate multiple-choice Algebra questions covering the topic and standards above.
2. STRICT RATIONAL ANSWERS ONLY: Every numeric answer MUST be a clean rational number (an integer, an exact simple fraction like 3/4 or -1/2, or a short exact terminating decimal like 1.5 or 0.75). NEVER generate irrational numbers (such as √2 or π) or non-terminating/messy rounded decimals!
3. Double-check every single calculation! The choice at correctIndex MUST be mathematically 100% exact, and the other 3 distractors must be plausible common algebra mistakes (sign error, reciprocal, forgot constant, etc.) and NEVER equal to the correct answer.
4. Keep question prompts clear and readable on a game modal: use clean Unicode math like x², √, ≤, ≥, ≠, ÷, ·, and write fractions cleanly as 3/4, -1/2, or (3x - 1) / 2 (NO raw $$ LaTeX syntax — the game automatically displays typed fractions vertically). Put equations, inequalities, systems, and multi-line steps on their own lines using \n, and separate paired system equations in answer choices with "  and  ".
5. Each question must be solvable without needing an external picture unless all coordinates/equations are explicitly stated in the prompt!
6. Provide a clear 1-sentence explanation showing how to get the answer.
7. Mix difficulties across easy, medium, and hard.`;

      const parts: Part[] = [...pdfParts, { text: textPrompt }];

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: { parts },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              questions: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    prompt: { type: Type.STRING },
                    choices: {
                      type: Type.ARRAY,
                      items: { type: Type.STRING },
                    },
                    correctIndex: { type: Type.INTEGER },
                    explanation: { type: Type.STRING },
                    standardCode: { type: Type.STRING },
                    difficulty: { type: Type.STRING },
                  },
                  required: ['prompt', 'choices', 'correctIndex', 'explanation', 'standardCode', 'difficulty'],
                },
              },
            },
            required: ['questions'],
          },
        },
      });

      const parsed = JSON.parse((response.text || '{"questions":[]}').trim());
      const rawQs = Array.isArray(parsed.questions) ? parsed.questions : [];

      const seenNormPrompts = new Set<string>(
        avoidList.map((p: string) => p.toLowerCase().replace(/[^a-z0-9]/g, ''))
      );

      const cleanQuestions = [];
      for (let i = 0; i < rawQs.length; i++) {
        const q = rawQs[i];
        const prompt = String(q.prompt || '').trim();
        if (!prompt) continue;
        const normKey = prompt.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (seenNormPrompts.has(normKey)) continue;
        seenNormPrompts.add(normKey);

        const choices = Array.isArray(q.choices)
          ? q.choices.map((c: any) => String(c).trim()).filter(Boolean).slice(0, 4)
          : [];
        if (choices.length !== 4) continue;
        const correctIndex = Math.max(0, Math.min(3, Math.trunc(Number(q.correctIndex) || 0)));
        const rawDiff = String(q.difficulty || 'medium').toLowerCase();
        const difficulty = rawDiff === 'easy' || rawDiff === 'hard' ? rawDiff : 'medium';

        cleanQuestions.push({
          id: `q_${Date.now().toString(36)}_${i}`,
          prompt: prompt.slice(0, 500),
          choices: choices.map((c: string) => c.slice(0, 180)),
          correctIndex,
          explanation: String(q.explanation || 'Check each step carefully.').trim().slice(0, 320),
          standardCode: String(q.standardCode || stdList[0] || '').trim().slice(0, 120),
          difficulty,
        });
      }

      res.json({
        questions: cleanQuestions,
        resourceReport,
      });
    } catch (err: any) {
      res.status(500).json({
        error: err?.message || 'Failed to generate questions from Gemini.',
        resourceReport,
      });
    }
  });

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.resolve(__dirname, 'dist');
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
      app.get('*', (_req: Request, res: Response) => {
        res.sendFile(path.join(distPath, 'index.html'));
      });
    }
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Algebra Dynasty '92 server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Fatal server startup error:', err);
  process.exit(1);
});
