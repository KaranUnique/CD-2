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
    { id: 'pa',  name: 'PANews',          url: 'https://www.panewslab.com/rss'                   },
    { id: 'cc',  name: 'ChainCatcher',    url: 'https://www.chaincatcher.com/rss'                },
    { id: 'tf',  name: 'TechFlow',        url: 'https://www.techflowpost.com/rss'                },
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
      const xml   = fetchFeed_(src.url);
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

    const limit = Math.min(
      parseInt(params.limit, 10) || CONFIG.DEFAULT_LIMIT,
      1000
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
//  PRIVATE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

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
      source     : String(row[COL.SOURCE]      || ''),
      pubDate    : row[COL.PUB_DATE] ? new Date(row[COL.PUB_DATE]).toISOString() : '',
    }));
}

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
      const title   = getText('title');
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
        description: description.substring(0, 500),
        link       : link,
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