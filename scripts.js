/**
 * ============================================================
 *  CryptoDesk — Google Apps Script Backend
 *  File: Code.gs
 *
 *  Responsibilities:
 *    1. fetchAndStoreRSS()  — Collect RSS feeds → Google Sheet
 *    2. doGet(e)            — Web API endpoint → JSON for frontend
 *    3. setupTrigger()      — One-time trigger installation helper
 * ============================================================
 */

// ── CONFIGURATION ─────────────────────────────────────────────────────────────

const CONFIG = {
  SHEET_NAME       : 'News',
  TRIGGER_INTERVAL : 30,          // minutes between auto-syncs (set in setupTrigger)
  MAX_ITEMS_PER_FEED: 50,         // max articles to pull per feed per run
  DEFAULT_LIMIT    : 200,         // default API response limit

  RSS_SOURCES: [
    { id: 'ct',  name: 'CoinTelegraph',   url: 'https://cointelegraph.com/rss'                   },
    { id: 'cd',  name: 'CoinDesk',        url: 'https://www.coindesk.com/arc/outboundfeeds/rss/' },
    { id: 'dc',  name: 'Decrypt',         url: 'https://decrypt.co/feed'                         },
    { id: 'cb',  name: 'Crypto Briefing', url: 'https://cryptobriefing.com/feed/'                },
    { id: 'bm',  name: 'Bitcoin News',    url: 'https://news.bitcoin.com/feed/'                  },
  ],
};

// ── SHEET COLUMN MAP (0-indexed) ──────────────────────────────────────────────

const COL = {
  GUID        : 0,
  TITLE       : 1,
  DESCRIPTION : 2,
  SOURCE      : 3,
  LINK        : 4,
  PUB_DATE    : 5,
  FETCH_TS    : 6,
};

const HEADERS = ['GUID', 'Title', 'Description', 'Source', 'Link', 'Published Date', 'Fetch Timestamp'];

// ─────────────────────────────────────────────────────────────────────────────
//  1.  RSS COLLECTOR  —  fetchAndStoreRSS()
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Main entry point.
 * Fetches all configured RSS feeds and appends only new articles to the Sheet.
 * Safe to run manually or via time-driven trigger.
 */
function fetchAndStoreRSS() {
  const sheet     = getOrCreateSheet_();
  const existingGuids = loadExistingGuids_(sheet);   // Set<string>
  const now       = new Date().toISOString();
  const rowsToAdd = [];
  const errors    = [];

  Logger.log('▶ Starting RSS sync. Existing records: %s', existingGuids.size);

  CONFIG.RSS_SOURCES.forEach(src => {
    try {
      const xml   = fetchFeed_(src.url);
      const items = parseRSS_(xml, src.name);
      let added   = 0;

      items.forEach(item => {
        if (existingGuids.has(item.guid)) return;   // already in sheet
        existingGuids.add(item.guid);               // prevent same-run dupes
        rowsToAdd.push([
          item.guid,
          item.title,
          item.description,
          item.source,
          item.link,
          item.pubDate,
          now,
        ]);
        added++;
      });

      Logger.log('  ✓ %s → %s new articles (of %s parsed)', src.name, added, items.length);
    } catch (err) {
      errors.push({ source: src.name, error: err.message });
      Logger.log('  ✗ %s → ERROR: %s', src.name, err.message);
    }
  });

  // Batch-append all new rows in a single sheet operation
  if (rowsToAdd.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rowsToAdd.length, HEADERS.length)
         .setValues(rowsToAdd);
    Logger.log('✔ Appended %s new rows to sheet.', rowsToAdd.length);
  } else {
    Logger.log('✔ No new articles found this run.');
  }

  if (errors.length) {
    Logger.log('⚠ Errors encountered: %s', JSON.stringify(errors));
  }

  return { inserted: rowsToAdd.length, errors };
}

// ─────────────────────────────────────────────────────────────────────────────
//  2.  WEB API  —  doGet(e)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * HTTP GET handler — deployed as a Web App.
 *
 * Query parameters (all optional):
 *   ?limit=50          max rows returned         (default: CONFIG.DEFAULT_LIMIT)
 *   ?source=CoinDesk   filter by source name      (case-insensitive, partial match)
 *   ?search=bitcoin    full-text search on title  (case-insensitive)
 *
 * Response: JSON array sorted by Published Date descending.
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  try {
    const sheet = getOrCreateSheet_();
    const data  = readSheetData_(sheet);

    // ── Filters ───────────────────────────────────────────────────────────────
    let results = data;

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

    // ── Sort: newest first ────────────────────────────────────────────────────
    results.sort((a, b) => {
      const da = a.pubDate ? new Date(a.pubDate).getTime() : 0;
      const db = b.pubDate ? new Date(b.pubDate).getTime() : 0;
      return db - da;
    });

    // ── Limit ─────────────────────────────────────────────────────────────────
    const limit = Math.min(
      parseInt(params.limit, 10) || CONFIG.DEFAULT_LIMIT,
      1000   // hard cap — protect against abuse
    );
    results = results.slice(0, limit);

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
//  3.  TRIGGER SETUP  —  run once from the Script Editor
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Call this function ONCE from the Apps Script editor to install a
 * time-driven trigger.  It removes any existing triggers first so
 * it's idempotent.
 *
 * How to run:
 *   1. Open script editor → select "setupTrigger" from the function dropdown
 *   2. Click ▶ Run
 *   3. Accept any permission prompts
 */
function setupTrigger() {
  // Remove existing triggers for fetchAndStoreRSS to avoid duplicates
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
//  PRIVATE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns the "News" sheet, creating it (with headers) if it doesn't exist.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold')
               .setBackground('#1a1a2e')
               .setFontColor('#ffffff');

    // Freeze header row
    sheet.setFrozenRows(1);

    // Set column widths
    sheet.setColumnWidth(1, 240);  // GUID
    sheet.setColumnWidth(2, 340);  // Title
    sheet.setColumnWidth(3, 400);  // Description
    sheet.setColumnWidth(4, 140);  // Source
    sheet.setColumnWidth(5, 260);  // Link
    sheet.setColumnWidth(6, 180);  // Published Date
    sheet.setColumnWidth(7, 180);  // Fetch Timestamp

    Logger.log('Created new sheet: %s', CONFIG.SHEET_NAME);
  }

  return sheet;
}

/**
 * Loads all existing GUIDs from the sheet into a Set for O(1) lookup.
 * Only reads the GUID column — minimises data transfer.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Set<string>}
 */
function loadExistingGuids_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const guids = sheet.getRange(2, COL.GUID + 1, lastRow - 1, 1).getValues();
  return new Set(guids.flat().filter(Boolean).map(String));
}

/**
 * Reads ALL data rows from the sheet and returns an array of article objects.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<Object>}
 */
function readSheetData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return values
    .filter(row => row[COL.GUID])   // skip blank rows
    .map(row => ({
      guid       : String(row[COL.GUID]        || ''),
      title      : String(row[COL.TITLE]       || ''),
      description: String(row[COL.DESCRIPTION] || ''),
      source     : String(row[COL.SOURCE]      || ''),
      link       : String(row[COL.LINK]        || ''),
      pubDate    : row[COL.PUB_DATE] ? new Date(row[COL.PUB_DATE]).toISOString() : '',
      fetchTs    : row[COL.FETCH_TS] ? new Date(row[COL.FETCH_TS]).toISOString() : '',
    }));
}

/**
 * Fetches raw XML from an RSS feed URL.
 * @param {string} url
 * @returns {string} raw XML text
 */
function fetchFeed_(url) {
  const options = {
    muteHttpExceptions: true,
    followRedirects   : true,
    headers           : {
      'User-Agent': 'Mozilla/5.0 (compatible; CryptoDeskBot/1.0)',
      'Accept'    : 'application/rss+xml, application/xml, text/xml, */*',
    },
  };

  const response = UrlFetchApp.fetch(url, options);
  const code     = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(`HTTP ${code} from ${url}`);
  }

  return response.getContentText();
}

/**
 * Parses an RSS XML string into an array of article objects.
 * Handles RSS 2.0 and Atom feeds.
 * @param {string} xmlText
 * @param {string} sourceName
 * @returns {Array<{guid, title, description, source, link, pubDate}>}
 */
function parseRSS_(xmlText, sourceName) {
  let doc;
  try {
    doc = XmlService.parse(xmlText);
  } catch (e) {
    throw new Error(`XML parse failed: ${e.message}`);
  }

  const root    = doc.getRootElement();
  const ns      = root.getNamespace();
  const dcNs    = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');
  const contentNs = XmlService.getNamespace('content', 'http://purl.org/rss/1.0/modules/content/');
  const isAtom  = root.getName() === 'feed';

  const items = [];

  if (isAtom) {
    // ── Atom feed ──────────────────────────────────────────────────────────
    root.getChildren('entry', ns).slice(0, CONFIG.MAX_ITEMS_PER_FEED).forEach(entry => {
      const getText = tag => safeGetText_(entry, tag, ns);

      const idEl   = entry.getChild('id', ns);
      const guid   = idEl ? idEl.getText().trim() : '';

      const linkEl = entry.getChild('link', ns);
      const link   = linkEl ? (linkEl.getAttributeValue('href') || '') : '';

      const title  = getText('title');

      const summaryEl  = entry.getChild('summary', ns);
      const contentEl  = entry.getChild('content', ns);
      const description = (summaryEl || contentEl)
        ? stripHtml_((summaryEl || contentEl).getText())
        : '';

      const pubDate = getText('published') || getText('updated');

      if (!title && !link) return;

      items.push({
        guid       : guid || link,
        title      : title || 'Untitled',
        description: description.substring(0, 500),
        source     : sourceName,
        link       : link || guid,
        pubDate    : pubDate ? new Date(pubDate).toISOString() : '',
      });
    });

  } else {
    // ── RSS 2.0 feed ───────────────────────────────────────────────────────
    const channel = root.getChild('channel') || root;
    channel.getChildren('item').slice(0, CONFIG.MAX_ITEMS_PER_FEED).forEach(item => {
      const getText = tag => safeGetText_(item, tag, null);

      const title       = getText('title');
      const link        = getText('link') || getText('guid');
      const guidEl      = item.getChild('guid');
      const guid        = guidEl ? guidEl.getText().trim() : link;

      // Description: prefer content:encoded → description
      let description = '';
      try {
        const contentEncoded = item.getChild('encoded', contentNs);
        if (contentEncoded) description = stripHtml_(contentEncoded.getText());
      } catch(_) {}
      if (!description) description = stripHtml_(getText('description'));

      // Published date: try dc:date, pubDate, date
      let pubDate = getText('pubDate');
      if (!pubDate) {
        try {
          const dcDate = item.getChild('date', dcNs);
          if (dcDate) pubDate = dcDate.getText().trim();
        } catch(_) {}
      }

      if (!title && !link) return;

      items.push({
        guid       : guid || link,
        title      : title || 'Untitled',
        description: description.substring(0, 500),
        source     : sourceName,
        link       : link || guid,
        pubDate    : pubDate ? safeParseDate_(pubDate) : '',
      });
    });
  }

  return items;
}

/**
 * Safely get text content of a child element.
 */
function safeGetText_(parent, tag, ns) {
  try {
    const child = ns ? parent.getChild(tag, ns) : parent.getChild(tag);
    return child ? child.getText().trim() : '';
  } catch (_) {
    return '';
  }
}

/**
 * Parse date string safely — returns ISO string or empty string.
 */
function safeParseDate_(str) {
  try {
    const d = new Date(str);
    return isNaN(d.getTime()) ? '' : d.toISOString();
  } catch (_) {
    return '';
  }
}

/**
 * Strip HTML tags and decode common entities from a string.
 * @param {string} html
 * @returns {string}
 */
function stripHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g,   '&')
    .replace(/&lt;/g,    '<')
    .replace(/&gt;/g,    '>')
    .replace(/&quot;/g,  '"')
    .replace(/&#39;/g,   "'")
    .replace(/&nbsp;/g,  ' ')
    .replace(/\s+/g,     ' ')
    .trim();
}

/**
 * Build a JSON ContentService response with CORS headers.
 * @param {Object} payload
 * @param {number} [status]
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function buildResponse_(payload, status) {
  const output = ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}