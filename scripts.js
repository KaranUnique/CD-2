/**
 * ============================================================
 *  CryptoDesk — Google Apps Script Backend
 *  File: Code.gs
 *
 *  Columns: S.No | Title | Description | Link | Source | Published Date
 * ============================================================
 */

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

const CONFIG = {
  SHEET_NAME        : 'News',
  TRIGGER_INTERVAL  : 30,
  MAX_ITEMS_PER_FEED: 50,
  DEFAULT_LIMIT     : 200,

  RSS_SOURCES: [
    { id: 'ct',  name: 'CoinTelegraph',   url: 'https://cointelegraph.com/rss'                   },
    { id: 'cd',  name: 'CoinDesk',        url: 'https://www.coindesk.com/arc/outboundfeeds/rss/' },
    { id: 'dc',  name: 'Decrypt',         url: 'https://decrypt.co/feed'                         },
    { id: 'cb',  name: 'Crypto Briefing', url: 'https://cryptobriefing.com/feed/'                },
    { id: 'bm',  name: 'Bitcoin News',    url: 'https://news.bitcoin.com/feed/'                  },
    { id: 'pa',  name: 'PANews',          url: 'https://www.panewslab.com/rss.xml'              },
    { id: 'cc',  name: 'ChainCatcher',    url: 'https://news.google.com/rss/search?q=site:chaincatcher.com&hl=en-US&gl=US&ceid=US:en' },
    { id: 'tf',  name: 'TechFlow',        url: 'https://news.google.com/rss/search?q=TechFlow+crypto&hl=en-US&gl=US&ceid=US:en'         },
  ],
};

// ── SHEET COLUMN MAP (0-indexed) ──────────────────────────────────────────────

const COL = {
  SNO         : 0,   // A
  TITLE       : 1,   // B
  DESCRIPTION : 2,   // C
  LINK        : 3,   // D
  SOURCE      : 4,   // E
  PUB_DATE    : 5,   // F
};

const HEADERS = ['S.No', 'Title', 'Description', 'Link', 'Source', 'Published Date'];

// ─────────────────────────────────────────────────────────────────────────────
//  1.  RSS COLLECTOR  —  fetchAndStoreRSS()
// ─────────────────────────────────────────────────────────────────────────────

function fetchAndStoreRSS() {
  const sheet        = getOrCreateSheet_();
  const existingLinks = loadExistingLinks_(sheet);   // Set<string> — dedupe by link
  const rowsToAdd    = [];
  const errors       = [];

  // Current last row to calculate starting S.No
  let nextSNo = sheet.getLastRow();   // header is row 1, so lastRow = count of data rows + 1

  Logger.log('▶ Starting RSS sync. Existing records: %s', existingLinks.size);

  CONFIG.RSS_SOURCES.forEach(src => {
    try {
      const xml   = fetchFeed_(src.url, src.name);
      const items = parseRSS_(xml, src.name);
      let added   = 0;

      items.forEach(item => {
        const key = item.link || item.title;
        if (existingLinks.has(key)) return;
        existingLinks.add(key);
        nextSNo++;
        rowsToAdd.push([
          nextSNo,
          item.title,
          item.description,
          item.link,
          item.source,
          item.pubDate,
        ]);
        added++;
      });

      Logger.log('  ✓ %s → %s new articles (of %s parsed)', src.name, added, items.length);
    } catch (err) {
      errors.push({ source: src.name, error: err.message });
      Logger.log('  ✗ %s → ERROR: %s', src.name, err.message);
    }
  });

  if (rowsToAdd.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAdd.length, HEADERS.length)
         .setValues(rowsToAdd);
    Logger.log('✔ Appended %s new rows.', rowsToAdd.length);
  } else {
    Logger.log('✔ No new articles this run.');
  }

  if (errors.length) {
    Logger.log('⚠ Errors: %s', JSON.stringify(errors));
  }

  return { inserted: rowsToAdd.length, errors };
}

// ─────────────────────────────────────────────────────────────────────────────
//  2.  WEB API  —  doGet(e)
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  try {
    const sheet = getOrCreateSheet_();
    const data  = readSheetData_(sheet);

    let results = data;

    // Filter out rows with unknown/unresolved sources
    results = results.filter(r => r.source && r.source !== 'Unknown' && !/^https?:\/\//i.test(r.source));

    if (params.source) {
      const src = params.source.toLowerCase();
      results = results.filter(r => r.source.toLowerCase().includes(src));
    }

    if (params.search) {
      const q = params.search.toLowerCase();
      results = results.filter(r =>
        r.title.toLowerCase().includes(q) ||
        r.description.toLowerCase().includes(q)
      );
    }

    // Sort newest first
    results.sort((a, b) => {
      const da = a.pubDate ? new Date(a.pubDate).getTime() : 0;
      const db = b.pubDate ? new Date(b.pubDate).getTime() : 0;
      return db - da;
    });

    // Apply per-source limit so no single source dominates when limit is hit
    const limit = Math.min(
      parseInt(params.limit, 10) || CONFIG.DEFAULT_LIMIT,
      1000
    );

    // Group by source, take top N per source, then flatten sorted
    const perSourceLimit = Math.ceil(limit / CONFIG.RSS_SOURCES.length);
    const bySource = {};
    results.forEach(r => {
      if (!bySource[r.source]) bySource[r.source] = [];
      if (bySource[r.source].length < perSourceLimit) bySource[r.source].push(r);
    });
    results = Object.values(bySource).flat()
      .sort((a, b) => {
        const da = a.pubDate ? new Date(a.pubDate).getTime() : 0;
        const db = b.pubDate ? new Date(b.pubDate).getTime() : 0;
        return db - da;
      })
      .slice(0, limit);

    const payload = {
      status  : 'ok',
      count   : results.length,
      total   : data.length,
      fetched : new Date().toISOString(),
      articles: results,
    };

    return buildResponse_(payload);

  } catch (err) {
    Logger.log('doGet error: %s', err.message);
    return buildResponse_({ status: 'error', message: err.message }, 500);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  3.  TRIGGER SETUP
// ─────────────────────────────────────────────────────────────────────────────

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fetchAndStoreRSS') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('fetchAndStoreRSS')
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL)
    .create();

  Logger.log('✔ Trigger created: fetchAndStoreRSS every %s minutes.', CONFIG.TRIGGER_INTERVAL);
}

// ─────────────────────────────────────────────────────────────────────────────
//  ONE-TIME SHEET REPAIR  —  fixSheetSources()
//  Run this manually once from the Apps Script editor to clean up dirty rows.
//  Fetches live Google News feeds to build a link→source map, then patches sheet.
// ─────────────────────────────────────────────────────────────────────────────

function fixSheetSources() {
  const sheet   = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nothing to fix.'); return; }

  const range  = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const values = range.getValues();

  // Build link→source map from live Google News feeds
  const googleNewsSources = [
    { name: 'ChainCatcher', url: 'https://news.google.com/rss/search?q=site:chaincatcher.com&hl=en-US&gl=US&ceid=US:en' },
    { name: 'TechFlow',     url: 'https://news.google.com/rss/search?q=TechFlow+crypto&hl=en-US&gl=US&ceid=US:en'     },
    { name: 'PANews',       url: 'https://news.google.com/rss/search?q=site:panewslab.com&hl=en-US&gl=US&ceid=US:en'  },
  ];

  // linkSuffix (last part of Google redirect URL, unique per article) → source name
  const linkSuffixMap = {};

  googleNewsSources.forEach(src => {
    try {
      const xml   = fetchUrl_(src.url);
      const doc   = XmlService.parse(xml);
      const channel = doc.getRootElement().getChild('channel') || doc.getRootElement();
      channel.getChildren('item').forEach(item => {
        const link = safeGetText_(item, 'link', null) || safeGetText_(item, 'guid', null);
        if (link) {
          // Store full link and also a suffix key (last 30 chars) for fuzzy matching
          linkSuffixMap[link] = src.name;
          linkSuffixMap[link.slice(-40)] = src.name;
        }
      });
      Logger.log('  Mapped %s links for %s', Object.keys(linkSuffixMap).length, src.name);
    } catch (e) {
      Logger.log('  Could not fetch %s: %s', src.name, e.message);
    }
  });

  let fixed = 0;

  values.forEach((row, i) => {
    const src  = String(row[COL.SOURCE] || '');
    const link = String(row[COL.LINK]   || '');

    // Skip rows that already have a clean, known source name
    if (src && !/^https?:\/\//i.test(src) && src !== 'Unknown' && src !== '') return;

    let newSrc = null;

    // 1. Direct domain lookup from link
    try {
      const hostname = new URL(link).hostname.replace(/^www\./, '');
      for (const [domain, name] of Object.entries(DOMAIN_TO_SOURCE)) {
        if (hostname.includes(domain.replace(/^www\./, ''))) { newSrc = name; break; }
      }
    } catch (_) {}

    // 2. Match against live Google News feed links
    if (!newSrc) {
      newSrc = linkSuffixMap[link] || linkSuffixMap[link.slice(-40)] || null;
    }

    // 3. If link is a Google News URL, assign based on S.No range heuristic
    //    (last resort — mark as needing manual review)
    if (!newSrc && link.includes('news.google.com')) {
      // We can't determine source — leave as-is but log for awareness
      Logger.log('  Cannot determine source for row %s: %s', i + 2, link.substring(0, 80));
    }

    if (newSrc && newSrc !== src) {
      values[i][COL.SOURCE] = newSrc;
      fixed++;
    }
  });

  if (fixed > 0) {
    range.setValues(values);
    Logger.log('✔ Fixed %s dirty source rows.', fixed);
  } else {
    Logger.log('✔ No dirty rows to fix (or could not match by live feeds).');
  }
}

/**
 * deleteUnknownRows()
 * Deletes all rows where Source is empty, a URL, or "Unknown".
 * Run this once manually, then run fetchAndStoreRSS to re-fetch clean data.
 */
function deleteUnknownRows() {
  const sheet   = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nothing to delete.'); return; }

  const values = sheet.getRange(2, COL.SOURCE + 1, lastRow - 1, 1).getValues();
  const rowsToDelete = [];

  values.forEach((row, i) => {
    const src = String(row[0] || '');
    if (!src || /^https?:\/\//i.test(src) || src === 'Unknown') {
      rowsToDelete.push(i + 2); // +2 for header row and 0-index offset
    }
  });

  // Delete from bottom to top to preserve row indices
  rowsToDelete.reverse().forEach(r => sheet.deleteRow(r));

  // Renumber S.No column
  const newLast = sheet.getLastRow();
  if (newLast >= 2) {
    const snoRange = sheet.getRange(2, COL.SNO + 1, newLast - 1, 1);
    const snoVals  = snoRange.getValues().map((_, i) => [i + 1]);
    snoRange.setValues(snoVals);
  }

  Logger.log('✔ Deleted %s unknown/dirty rows. Run fetchAndStoreRSS to re-fetch.', rowsToDelete.length);
}

function getOrCreateSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1a1a2e')
               .setFontColor('#ffffff')
               .setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // Column widths
    sheet.setColumnWidth(1, 60);   // S.No
    sheet.setColumnWidth(2, 340);  // Title
    sheet.setColumnWidth(3, 400);  // Description
    sheet.setColumnWidth(4, 260);  // Link
    sheet.setColumnWidth(5, 140);  // Source
    sheet.setColumnWidth(6, 180);  // Published Date

    Logger.log('Created new sheet: %s', CONFIG.SHEET_NAME);
  }

  return sheet;
}

/**
 * Load existing article links for deduplication (replaces GUID-based check).
 */
function loadExistingLinks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const links = sheet.getRange(2, COL.LINK + 1, lastRow - 1, 1).getValues();
  return new Set(links.flat().filter(Boolean).map(String));
}

// Map of link domains → canonical source names (for fixing dirty rows)
const DOMAIN_TO_SOURCE = {
  'cointelegraph.com'   : 'CoinTelegraph',
  'coindesk.com'        : 'CoinDesk',
  'decrypt.co'          : 'Decrypt',
  'cryptobriefing.com'  : 'Crypto Briefing',
  'news.bitcoin.com'    : 'Bitcoin News',
  'panewslab.com'       : 'PANews',
  'chaincatcher.com'    : 'ChainCatcher',
  'techflowpost.com'    : 'TechFlow',
};

function inferSource_(source, link, title, description) {
  // If source is already a valid name (not a URL), trust it
  if (source && !/^https?:\/\//i.test(source)) return source;

  // Check for embedded [SRC:name] tag in description (added by new fetches)
  const srcTag = (description || '').match(/\[SRC:([^\]]+)\]/);
  if (srcTag) return srcTag[1];

  // Try to derive from the link domain
  try {
    const hostname = new URL(link).hostname.replace(/^www\./, '');
    for (const [domain, name] of Object.entries(DOMAIN_TO_SOURCE)) {
      if (hostname.includes(domain.replace(/^www\./, ''))) return name;
    }
  } catch (_) {}

  return 'Unknown';
}

function readSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return values
    .filter(row => row[COL.TITLE])   // skip blank rows
    .map(row => ({
      sno        : row[COL.SNO]         || '',
      title      : String(row[COL.TITLE]       || ''),
      description: String(row[COL.DESCRIPTION] || ''),
      link       : String(row[COL.LINK]        || ''),
      source     : inferSource_(String(row[COL.SOURCE] || ''), String(row[COL.LINK] || ''), String(row[COL.TITLE] || ''), String(row[COL.DESCRIPTION] || '')),
      pubDate    : row[COL.PUB_DATE] ? new Date(row[COL.PUB_DATE]).toISOString() : '',
    }));
}

// Known fallback URLs for sources that have unstable feed paths
const FEED_FALLBACKS = {
  'PANews'      : ['https://www.panewslab.com/rss.xml', 'https://www.panewslab.com/zh/rss', 'https://www.panewslab.com/feed'],
  'ChainCatcher': [
    'https://news.google.com/rss/search?q=site:chaincatcher.com&hl=en-US&gl=US&ceid=US:en',
    'https://news.google.com/rss/search?q=ChainCatcher+crypto&hl=en-US&gl=US&ceid=US:en',
  ],
  'TechFlow'    : [
    'https://news.google.com/rss/search?q=TechFlow+crypto&hl=en-US&gl=US&ceid=US:en',
    'https://news.google.com/rss/search?q=techflowpost.com&hl=en-US&gl=US&ceid=US:en',
  ],
};

/**
 * Google News RSS wraps article links in a redirect URL.
 * Attempt to resolve it so the stored link points to the real article.
 * Falls back to the redirect URL if resolution fails.
 */
function resolveGoogleNewsLink_(redirectUrl) {
  if (!redirectUrl || !redirectUrl.includes('news.google.com')) return redirectUrl;
  try {
    const options = {
      muteHttpExceptions: true,
      followRedirects   : false,
      headers           : { 'User-Agent': 'Mozilla/5.0 (compatible; CryptoDeskBot/1.0)' },
    };
    const r = UrlFetchApp.fetch(redirectUrl, options);
    const code = r.getResponseCode();
    if (code >= 300 && code < 400) {
      const loc = r.getHeaders()['Location'] || r.getHeaders()['location'];
      if (loc) return loc;
    }
    // Try extracting URL from the response body (Google sometimes embeds it)
    const body = r.getContentText().substring(0, 2000);
    const m = body.match(/url=(https?:\/\/[^&"'\s]+)/i);
    if (m) return decodeURIComponent(m[1]);
  } catch (_) {}
  return redirectUrl;
}

function fetchFeed_(url, sourceName) {
  const urls = (sourceName && FEED_FALLBACKS[sourceName]) ? FEED_FALLBACKS[sourceName] : [url];

  let lastErr = '';
  for (const tryUrl of urls) {
    try {
      const text = fetchUrl_(tryUrl);
      return text;
    } catch (e) {
      lastErr = e.message;
      Logger.log('  ↳ tried %s → %s', tryUrl, e.message);
    }
  }
  throw new Error(lastErr);
}

function fetchUrl_(url) {
  const options = {
    muteHttpExceptions: true,
    followRedirects   : true,
    headers           : {
      'User-Agent'     : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept'         : 'application/rss+xml, application/atom+xml, application/xml, text/xml, */*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
      'Cache-Control'  : 'no-cache',
    },
  };

  const response = UrlFetchApp.fetch(url, options);
  const code     = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(`HTTP ${code} from ${url}`);
  }

  const text = response.getContentText();

  // Reject HTML pages masquerading as feeds
  if (/^\s*<!DOCTYPE\s+html/i.test(text) || /<html[\s>]/i.test(text.substring(0, 500))) {
    throw new Error(`URL returned an HTML page, not an RSS/XML feed`);
  }

  return sanitizeXml_(text);
}

/**
 * Fix common XML issues in RSS feeds before strict parsing:
 * - Unescaped bare & that are not part of a valid entity reference
 */
function sanitizeXml_(xml) {
  // Replace bare & not followed by a valid entity or numeric ref with &amp;
  return xml.replace(/&(?!(?:#\d+|#x[\da-fA-F]+|amp|lt|gt|quot|apos|nbsp);)/g, '&amp;');
}

function parseRSS_(xmlText, sourceName) {
  let doc;
  try {
    doc = XmlService.parse(xmlText);
  } catch (e) {
    throw new Error(`XML parse failed: ${e.message}`);
  }

  const root      = doc.getRootElement();
  const ns        = root.getNamespace();
  const dcNs      = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');
  const contentNs = XmlService.getNamespace('content', 'http://purl.org/rss/1.0/modules/content/');
  const isAtom    = root.getName() === 'feed';
  const items     = [];

  if (isAtom) {
    root.getChildren('entry', ns).slice(0, CONFIG.MAX_ITEMS_PER_FEED).forEach(entry => {
      const getText = tag => safeGetText_(entry, tag, ns);
      const idEl   = entry.getChild('id', ns);
      const linkEl = entry.getChild('link', ns);
      const link   = linkEl ? (linkEl.getAttributeValue('href') || '') : '';
      const title  = getText('title');
      const summaryEl  = entry.getChild('summary', ns);
      const contentEl  = entry.getChild('content', ns);
      const description = (summaryEl || contentEl)
        ? stripHtml_((summaryEl || contentEl).getText()) : '';
      const pubDate = getText('published') || getText('updated');
      if (!title && !link) return;
      items.push({
        title      : title || 'Untitled',
        description: description.substring(0, 500),
        link       : link || (idEl ? idEl.getText().trim() : ''),
        source     : sourceName,
        pubDate    : pubDate ? new Date(pubDate).toISOString() : '',
      });
    });

  } else {
    const channel = root.getChild('channel') || root;
    channel.getChildren('item').slice(0, CONFIG.MAX_ITEMS_PER_FEED).forEach(item => {
      const getText = tag => safeGetText_(item, tag, null);
      let   title   = getText('title');
      // Google News appends " - Source Name" suffix — strip it for Google News-proxied sources
      if (sourceName === 'ChainCatcher') {
        title = title.replace(/\s*-\s*(?:ChainCatcher|链捕手ChainCatcher).*$/i, '').trim();
      } else if (sourceName === 'TechFlow') {
        title = title.replace(/\s*-\s*(?:深潮TechFlow|TechFlow).*$/i, '').trim();
      }
      const link    = getText('link') || getText('guid');
      let description = '';
      try {
        const ce = item.getChild('encoded', contentNs);
        if (ce) description = stripHtml_(ce.getText());
      } catch(_) {}
      if (!description) description = stripHtml_(getText('description'));
      let pubDate = getText('pubDate');
      if (!pubDate) {
        try {
          const dcDate = item.getChild('date', dcNs);
          if (dcDate) pubDate = dcDate.getText().trim();
        } catch(_) {}
      }
      if (!title && !link) return;
      items.push({
        title      : title || 'Untitled',
        description: (description.substring(0, 490) + (sourceName === 'ChainCatcher' || sourceName === 'TechFlow' || sourceName === 'PANews' ? ` [SRC:${sourceName}]` : '')).trim(),
        link       : (sourceName === 'ChainCatcher' || sourceName === 'TechFlow') ? resolveGoogleNewsLink_(link) : link,
        source     : sourceName,
        pubDate    : pubDate ? safeParseDate_(pubDate) : '',
      });
    });
  }

  return items;
}

function safeGetText_(parent, tag, ns) {
  try {
    const child = ns ? parent.getChild(tag, ns) : parent.getChild(tag);
    return child ? child.getText().trim() : '';
  } catch (_) { return ''; }
}

function safeParseDate_(str) {
  try {
    const d = new Date(str);
    return isNaN(d.getTime()) ? '' : d.toISOString();
  } catch (_) { return ''; }
}

function stripHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function buildResponse_(payload, status) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}