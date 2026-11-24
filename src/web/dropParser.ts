export interface DroppedPayload {
  files: File[];
  candidates: string[];
  fallbackNames: string[];
  ignoredWebUrls: string[];
  rawTypes: string[];
}

const IGNORED_SCHEMES_RE = /^(?:https?|mailto|javascript|data|blob|about):/i;
const LINE_ANCHOR_SUFFIX_RE = /(?:#L\d+(?:-L?\d+)?|:\d+(?::\d+)?)$/i;

function unwrapQuotesAndBrackets(value: string, maxPasses = 2): string {
  let current = value.trim();
  const pairs: Array<[string, string]> = [
    ['"', '"'],
    ["'", "'"],
    ['`', '`'],
    ['<', '>'],
  ];
  for (let pass = 0; pass < maxPasses; pass++) {
    let changed = false;
    for (const [open, close] of pairs) {
      if (current.length >= 2 && current.startsWith(open) && current.endsWith(close)) {
        current = current.slice(open.length, current.length - close.length).trim();
        changed = true;
        break;
      }
    }
    if (!changed) break;
  }
  return current;
}

export function normalizePathCandidate(raw: string): string | null {
  let value = raw.trim();
  if (!value) return null;

  const mdMatch = value.match(/^\[[^\]]*\]\(([^)]+)\)$/);
  if (mdMatch) {
    value = mdMatch[1].trim();
  }

  value = unwrapQuotesAndBrackets(value, 2);
  if (!value) return null;

  if (value.startsWith('@')) {
    const afterAt = value.slice(1);
    if (/^(?:["'`]|file:\/\/|~[/\\]|\/|\.\.?[/\\]|[A-Za-z]:[/\\])/.test(afterAt)) {
      value = unwrapQuotesAndBrackets(afterAt, 2);
    }
  }

  if (!value || IGNORED_SCHEMES_RE.test(value)) {
    return null;
  }

  if (/^file:\/\//i.test(value)) {
    try {
      const u = new URL(value);
      let decoded = u.pathname;
      try {
        decoded = decodeURIComponent(u.pathname);
      } catch {
        // Keep raw pathname when percent-decoding fails
      }
      if (u.hostname && u.hostname.toLowerCase() !== 'localhost') {
        value = `\\\\${u.hostname}${decoded.replace(/\//g, '\\')}`;
      } else if (/^\/[A-Za-z]:/.test(decoded)) {
        value = decoded.slice(1).replace(/\//g, '\\');
      } else {
        value = decoded;
      }
    } catch {
      const stripped = value.replace(/^file:\/{2,3}/i, '').split(/[?#]/)[0];
      try {
        value = decodeURIComponent(stripped);
      } catch {
        value = stripped;
      }
      if (/^[A-Za-z]:\//.test(value)) {
        value = value.replace(/\//g, '\\');
      }
    }
  } else {
    value = value.replace(LINE_ANCHOR_SUFFIX_RE, '').trim();
  }

  return value || null;
}

export function looksLikeFilePath(raw: string): boolean {
  const s = raw.trim();
  if (!s) return false;
  if (IGNORED_SCHEMES_RE.test(s)) return false;
  if (/^file:\/\//i.test(s)) return true;
  if (/^[A-Za-z]:[/\\]/.test(s)) return true;
  if (/^\\\\[^\\]+\\[^\\]+/.test(s)) return true;
  if (/^~[/\\]/.test(s)) return true;
  if (/^\/[^\s<>|*?]+/.test(s)) return true;
  if (/^\.\.?[/\\][^\s<>|*?]+/.test(s)) return true;
  if (
    !s.includes('://') &&
    !/[<>*?|]/.test(s) &&
    /^[^\r\n]+[/\\][^\r\n]+\.[A-Za-z0-9]{1,8}(?:(?:#L\d+(?:-L?\d+)?|:\d+(?::\d+)?))?$/i.test(s)
  ) {
    return true;
  }
  return false;
}

export function looksLikeFilename(raw: string): boolean {
  let s = raw.trim();
  if (!s || /[\r\n]/.test(s)) return false;
  if (s.startsWith('@')) s = s.slice(1).trim();
  s = unwrapQuotesAndBrackets(s, 2);
  s = s.replace(LINE_ANCHOR_SUFFIX_RE, '').trim();
  if (s.length < 2 || s.length > 220) return false;
  if (s.includes('://') || /[/\\<>|*?"]/.test(s)) return false;
  return /\.[A-Za-z0-9]{1,8}$/.test(s);
}

export function extractPathsFromHtml(html: string): string[] {
  const results: string[] = [];
  if (!html) return results;

  try {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const pushAttr = (val: string | null, requirePathOrName = false) => {
      if (!val) return;
      const trimmed = val.trim();
      if (!trimmed) return;
      if (requirePathOrName && !looksLikeFilePath(trimmed) && !looksLikeFilename(trimmed)) return;
      results.push(trimmed);
    };

    doc.querySelectorAll('a[href]').forEach((el) => pushAttr(el.getAttribute('href')));
    doc.querySelectorAll('[data-uri]').forEach((el) => pushAttr(el.getAttribute('data-uri')));
    doc.querySelectorAll('[data-path]').forEach((el) => pushAttr(el.getAttribute('data-path')));
    doc.querySelectorAll('[data-file-path]').forEach((el) => pushAttr(el.getAttribute('data-file-path')));
    doc.querySelectorAll('[title]').forEach((el) => pushAttr(el.getAttribute('title'), true));
  } catch {
    // Ignore DOMParser errors and fall back to regex scan
  }

  const fileUriMatches = html.match(/file:\/\/\/[^\s"'<>]+/gi);
  if (fileUriMatches) {
    for (const uri of fileUriMatches) {
      results.push(uri);
    }
  }

  return results;
}

export function extractPathsFromJsonPayload(raw: string): string[] {
  const results: string[] = [];
  const trimmed = raw.trim();
  if (!trimmed || (!trimmed.startsWith('[') && !trimmed.startsWith('{'))) return results;

  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    return results;
  }

  const walk = (node: unknown, depth: number) => {
    if (depth > 6 || node == null) return;
    if (typeof node === 'string') {
      const s = node.trim();
      if (s) results.push(s);
      return;
    }
    if (Array.isArray(node)) {
      for (const item of node) walk(item, depth + 1);
      return;
    }
    if (typeof node === 'object') {
      const obj = node as Record<string, unknown>;
      if (obj.scheme === 'file' && (typeof obj.fsPath === 'string' || typeof obj.path === 'string')) {
        const picked = (typeof obj.fsPath === 'string' ? obj.fsPath : obj.path) as string;
        if (picked.trim()) results.push(picked.trim());
        return;
      }
      for (const key of ['uri', 'fsPath', 'filePath', 'path', 'url', 'resource', 'href', 'external']) {
        const val = obj[key];
        if (typeof val === 'string' && val.trim()) {
          results.push(val.trim());
        } else if (val && typeof val === 'object') {
          walk(val, depth + 1);
        }
      }
      for (const val of Object.values(obj)) {
        if (typeof val === 'string') {
          if (/^file:\/\//i.test(val.trim()) || looksLikeFilePath(val)) {
            results.push(val.trim());
          }
        } else if (val && typeof val === 'object') {
          walk(val, depth + 1);
        }
      }
    }
  };

  walk(parsed, 0);
  return results;
}

export function extractFromPlainText(raw: string): {
  candidates: string[];
  fallbackNames: string[];
  ignoredWebUrls: string[];
} {
  const candidates: string[] = [];
  const fallbackNames: string[] = [];
  const ignoredWebUrls: string[] = [];

  if (!raw) return { candidates, fallbackNames, ignoredWebUrls };

  const lines = raw.split(/\r?\n/);
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    // Web URLs recorded separately so user sees helpful message if only web links were dropped
    const webMatches = line.match(/https?:\/\/[^\s"'<>)\]]+/gi);
    if (webMatches) {
      for (const url of webMatches) ignoredWebUrls.push(url);
    }

    // Markdown links: [label](target)
    const mdGlobal = /\[([^\]]*)\]\(([^)]+)\)/g;
    let m: RegExpExecArray | null;
    let hadLineTarget = false;
    while ((m = mdGlobal.exec(line)) !== null) {
      const target = m[2].trim();
      if (/^https?:\/\//i.test(target)) {
        ignoredWebUrls.push(target);
        continue;
      }
      const norm = normalizePathCandidate(target);
      if (norm && (looksLikeFilePath(norm) || looksLikeFilePath(target))) {
        candidates.push(norm);
        hadLineTarget = true;
      } else if (norm && looksLikeFilename(norm)) {
        fallbackNames.push(norm.replace(/^@/, '').replace(LINE_ANCHOR_SUFFIX_RE, '').trim());
        hadLineTarget = true;
      }
    }

    // Explicit file:// URIs anywhere on the line
    const fileUriMatches = line.match(/file:\/\/[^\s"'<>)\]]+/gi);
    if (fileUriMatches) {
      for (const uri of fileUriMatches) {
        const norm = normalizePathCandidate(uri);
        if (norm) {
          candidates.push(norm);
          hadLineTarget = true;
        }
      }
    }

    // Backtick / double-quoted tokens
    const quotedRe = /`([^`]+)`|"([^"]+)"/g;
    while ((m = quotedRe.exec(line)) !== null) {
      const token = (m[1] ?? m[2] ?? '').trim();
      const norm = normalizePathCandidate(token);
      if (!norm) continue;
      if (looksLikeFilePath(norm) || looksLikeFilePath(token)) {
        candidates.push(norm);
        hadLineTarget = true;
      } else if (looksLikeFilename(token)) {
        fallbackNames.push(
          unwrapQuotesAndBrackets(token.replace(/^@/, ''), 2).replace(LINE_ANCHOR_SUFFIX_RE, '').trim(),
        );
        hadLineTarget = true;
      }
    }

    // Whole trimmed line as path or bare filename (handles unquoted Windows paths with spaces)
    if (!hadLineTarget) {
      const norm = normalizePathCandidate(line);
      if (norm && (looksLikeFilePath(norm) || looksLikeFilePath(line))) {
        candidates.push(norm);
      } else if (looksLikeFilename(line)) {
        fallbackNames.push(
          unwrapQuotesAndBrackets(line.replace(/^@/, ''), 2).replace(LINE_ANCHOR_SUFFIX_RE, '').trim(),
        );
      }
    }
  }

  return { candidates, fallbackNames, ignoredWebUrls };
}

export function extractDroppedPayload(dt: DataTransfer | null): DroppedPayload {
  const files = dt?.files ? Array.from(dt.files) : [];
  const rawTypes = dt?.types ? Array.from(dt.types) : [];
  const candidates: string[] = [];
  const fallbackNames: string[] = [];
  const ignoredWebUrls: string[] = [];

  if (!dt) {
    return { files, candidates, fallbackNames, ignoredWebUrls, rawTypes };
  }

  const seenCandidates = new Set<string>();
  const seenFallbacks = new Set<string>();
  const seenWeb = new Set<string>();

  const pushCandidate = (raw: string) => {
    const trimmed = raw.trim();
    if (!trimmed) return;
    if (/^https?:\/\//i.test(trimmed)) {
      if (!seenWeb.has(trimmed)) {
        seenWeb.add(trimmed);
        ignoredWebUrls.push(trimmed);
      }
      return;
    }
    const norm = normalizePathCandidate(trimmed);
    if (!norm) return;
    if (looksLikeFilePath(trimmed) || looksLikeFilePath(norm)) {
      if (!seenCandidates.has(norm)) {
        seenCandidates.add(norm);
        candidates.push(norm);
      }
    } else if (looksLikeFilename(trimmed) || looksLikeFilename(norm)) {
      const cleaned = unwrapQuotesAndBrackets(norm.replace(/^@/, ''), 2).replace(LINE_ANCHOR_SUFFIX_RE, '').trim();
      if (cleaned && !seenFallbacks.has(cleaned)) {
        seenFallbacks.add(cleaned);
        fallbackNames.push(cleaned);
      }
    }
  };

  const pushFallback = (name: string) => {
    const cleaned = unwrapQuotesAndBrackets(name.replace(/^@/, ''), 2).replace(LINE_ANCHOR_SUFFIX_RE, '').trim();
    if (cleaned && !seenFallbacks.has(cleaned)) {
      seenFallbacks.add(cleaned);
      fallbackNames.push(cleaned);
    }
  };

  const pushWeb = (url: string) => {
    if (url && !seenWeb.has(url)) {
      seenWeb.add(url);
      ignoredWebUrls.push(url);
    }
  };

  const safeGetData = (type: string): string => {
    try {
      return dt.getData(type) || '';
    } catch {
      return '';
    }
  };

  const processedTypes = new Set<string>(['Files', 'files']);

  // 1. text/uri-list
  const uriList = safeGetData('text/uri-list');
  processedTypes.add('text/uri-list');
  if (uriList) {
    for (const line of uriList.split(/\r?\n/)) {
      const t = line.trim();
      if (!t || t.startsWith('#')) continue;
      pushCandidate(t);
    }
  }

  // 2. VS Code / Electron editor payloads
  for (const structuredType of ['resourceurls', 'codefiles', 'application/vnd.code.uri-list']) {
    processedTypes.add(structuredType);
    const raw = safeGetData(structuredType);
    if (!raw) continue;
    const fromJson = extractPathsFromJsonPayload(raw);
    if (fromJson.length > 0) {
      for (const item of fromJson) pushCandidate(item);
    } else {
      for (const line of raw.split(/\r?\n/)) {
        const t = line.trim();
        if (t && !t.startsWith('#')) pushCandidate(t);
      }
    }
  }

  // 3. Mozilla URL list (alternating url / title lines)
  const mozUrl = safeGetData('text/x-moz-url');
  processedTypes.add('text/x-moz-url');
  if (mozUrl) {
    const lines = mozUrl.split(/\r?\n/);
    for (let i = 0; i < lines.length; i += 2) {
      if (lines[i]?.trim()) pushCandidate(lines[i].trim());
    }
  }

  // 4. Other custom MIME types mentioning uri / url / file / path / json
  for (const t of rawTypes) {
    if (processedTypes.has(t) || t === 'text/html' || t === 'text/plain') continue;
    if (!/(?:uri|url|file|path|json)/i.test(t)) continue;
    processedTypes.add(t);
    const raw = safeGetData(t);
    if (!raw) continue;
    const fromJson = extractPathsFromJsonPayload(raw);
    if (fromJson.length > 0) {
      for (const item of fromJson) pushCandidate(item);
    } else {
      const parsed = extractFromPlainText(raw);
      for (const c of parsed.candidates) pushCandidate(c);
      for (const f of parsed.fallbackNames) pushFallback(f);
      for (const w of parsed.ignoredWebUrls) pushWeb(w);
    }
  }

  // 5. text/html
  processedTypes.add('text/html');
  const html = safeGetData('text/html');
  if (html) {
    for (const item of extractPathsFromHtml(html)) {
      pushCandidate(item);
    }
  }

  // 6. text/plain
  processedTypes.add('text/plain');
  const plain = safeGetData('text/plain');
  if (plain) {
    const parsed = extractFromPlainText(plain);
    for (const c of parsed.candidates) pushCandidate(c);
    for (const f of parsed.fallbackNames) pushFallback(f);
    for (const w of parsed.ignoredWebUrls) pushWeb(w);
  }

  return { files, candidates, fallbackNames, ignoredWebUrls, rawTypes };
}
