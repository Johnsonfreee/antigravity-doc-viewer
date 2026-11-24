export interface ShadowHtmlDocument {
  resetCss: string;
  stylesCss: string;
  bodyHtml: string;
  htmlAttrs: Array<[string, string]>;
  bodyAttrs: Array<[string, string]>;
}

const DOC_VIEWER_CUSTOM_VARS = [
  '--chrome-tab-bg',
  '--chrome-addr-bg',
  '--chrome-tab-active',
  '--chrome-canvas',
  '--chrome-panel',
  '--chrome-panel-2',
  '--chrome-border',
  '--chrome-border-soft',
  '--chrome-hairline',
  '--chrome-fg',
  '--chrome-fg-2',
  '--chrome-fg-3',
  '--chrome-hover',
  '--chrome-hover-strong',
  '--chrome-addr-focus-bg',
  '--accent',
  '--accent-hover',
  '--accent-soft',
  '--accent-fg',
  '--bg-1',
  '--bg-2',
  '--bg-3',
  '--bg-info',
  '--bg-danger',
  '--fg-1',
  '--fg-2',
  '--fg-3',
  '--fg-info',
  '--fg-danger',
  '--bd-1',
  '--bd-2',
  '--wt-banner-bg',
  '--wt-banner-bd',
  '--wt-banner-fg',
  '--diff-del-bg',
  '--diff-del-fg',
  '--diff-add-bg',
  '--diff-add-fg',
  '--paper-shadow',
  '--popover-shadow',
  '--float-bg',
  '--float-shadow',
  '--backdrop-bg',
  '--badge-pdf-bg',
  '--badge-pdf-fg',
  '--badge-md-bg',
  '--badge-md-fg',
  '--badge-docx-bg',
  '--badge-docx-fg',
  '--badge-sheet-bg',
  '--badge-sheet-fg',
  '--badge-slides-bg',
  '--badge-slides-fg',
  '--badge-html-bg',
  '--badge-html-fg',
  '--badge-dir-bg',
  '--badge-dir-fg',
  '--font-ui',
  '--font-code',
  '--code-bg',
  '--code-fg',
  '--hl-kw',
  '--hl-str',
  '--hl-num',
  '--hl-fn',
  '--hl-com',
  '--hl-attr',
];

const VERBATIM_AT_RULES = new Set([
  '@keyframes',
  '@-webkit-keyframes',
  '@-moz-keyframes',
  '@font-face',
  '@page',
  '@property',
  '@font-feature-values',
  '@counter-style',
]);

const RECURSIVE_AT_RULES = new Set([
  '@supports',
  '@layer',
  '@container',
  '@scope',
  '@document',
  '@-moz-document',
]);

function stripCssComments(text: string): string {
  let out = '';
  let i = 0;
  let quote: string | null = null;
  while (i < text.length) {
    const ch = text[i];
    if (quote) {
      out += ch;
      if (ch === '\\' && i + 1 < text.length) {
        out += text[i + 1];
        i += 2;
        continue;
      }
      if (ch === quote) quote = null;
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      out += ch;
      i += 1;
      continue;
    }
    if (ch === '/' && text[i + 1] === '*') {
      const end = text.indexOf('*/', i + 2);
      i = end === -1 ? text.length : end + 2;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

function findMatchingBrace(css: string, openIdx: number): number {
  let depth = 0;
  let i = openIdx;
  let quote: string | null = null;
  while (i < css.length) {
    const ch = css[i];
    if (quote) {
      if (ch === '\\') {
        i += 2;
        continue;
      }
      if (ch === quote) quote = null;
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      i += 1;
      continue;
    }
    if (ch === '/' && css[i + 1] === '*') {
      const end = css.indexOf('*/', i + 2);
      i = end === -1 ? css.length : end + 2;
      continue;
    }
    if (ch === '{') depth += 1;
    else if (ch === '}') {
      depth -= 1;
      if (depth === 0) return i;
    }
    i += 1;
  }
  return css.length;
}

function splitTopLevelSelectors(prelude: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let parenDepth = 0;
  let bracketDepth = 0;
  let quote: string | null = null;
  let i = 0;

  while (i < prelude.length) {
    const ch = prelude[i];
    if (quote) {
      if (ch === '\\') {
        i += 2;
        continue;
      }
      if (ch === quote) quote = null;
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      i += 1;
      continue;
    }
    if (ch === '/' && prelude[i + 1] === '*') {
      const end = prelude.indexOf('*/', i + 2);
      i = end === -1 ? prelude.length : end + 2;
      continue;
    }
    if (ch === '(') parenDepth += 1;
    else if (ch === ')' && parenDepth > 0) parenDepth -= 1;
    else if (ch === '[') bracketDepth += 1;
    else if (ch === ']' && bracketDepth > 0) bracketDepth -= 1;
    else if (ch === ',' && parenDepth === 0 && bracketDepth === 0) {
      parts.push(prelude.slice(start, i));
      start = i + 1;
    }
    i += 1;
  }

  parts.push(prelude.slice(start));
  return parts;
}

function isIdentChar(ch: string): boolean {
  return /[A-Za-z0-9_-]/.test(ch);
}

function isValidTypePrefix(ch: string): boolean {
  return ch === '' || /[\s>+~(,|*/]/.test(ch);
}

function rewriteSingleSelector(rawSelector: string): string {
  const leadMatch = rawSelector.match(/^\s*/);
  const trailMatch = rawSelector.match(/\s*$/);
  const leading = leadMatch ? leadMatch[0] : '';
  const trailing = trailMatch && rawSelector.length > leading.length ? trailMatch[0] : '';
  const core = rawSelector.slice(leading.length, rawSelector.length - trailing.length);
  if (!core) return rawSelector;

  if (stripCssComments(core).trim() === ':root') {
    return `${leading}:host, .html-doc-html${trailing}`;
  }

  let out = '';
  let bracketDepth = 0;
  let quote: string | null = null;
  let i = 0;

  while (i < core.length) {
    const ch = core[i];
    if (quote) {
      out += ch;
      if (ch === '\\' && i + 1 < core.length) {
        out += core[i + 1];
        i += 2;
        continue;
      }
      if (ch === quote) quote = null;
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      out += ch;
      i += 1;
      continue;
    }
    if (ch === '/' && core[i + 1] === '*') {
      const end = core.indexOf('*/', i + 2);
      const sliceEnd = end === -1 ? core.length : end + 2;
      out += core.slice(i, sliceEnd);
      i = sliceEnd;
      continue;
    }
    if (ch === '[') {
      bracketDepth += 1;
      out += ch;
      i += 1;
      continue;
    }
    if (ch === ']' && bracketDepth > 0) {
      bracketDepth -= 1;
      out += ch;
      i += 1;
      continue;
    }

    if (bracketDepth === 0) {
      // Replace :root with .html-doc-html
      if (
        core.slice(i, i + 5).toLowerCase() === ':root' &&
        !isIdentChar(core[i + 5] || '')
      ) {
        out += '.html-doc-html';
        i += 5;
        continue;
      }

      // Replace standalone html type selector with .html-doc-html
      const prev = i === 0 ? '' : core[i - 1];
      if (
        core.slice(i, i + 4).toLowerCase() === 'html' &&
        isValidTypePrefix(prev) &&
        !isIdentChar(core[i + 4] || '') &&
        core[i + 4] !== '('
      ) {
        out += '.html-doc-html';
        i += 4;
        continue;
      }

      // Replace standalone body type selector with .html-doc-body
      if (
        core.slice(i, i + 4).toLowerCase() === 'body' &&
        isValidTypePrefix(prev) &&
        !isIdentChar(core[i + 4] || '') &&
        core[i + 4] !== '('
      ) {
        out += '.html-doc-body';
        i += 4;
        continue;
      }
    }

    out += ch;
    i += 1;
  }

  return `${leading}${out}${trailing}`;
}

function rewriteSelectorList(prelude: string): string {
  return splitTopLevelSelectors(prelude).map(rewriteSingleSelector).join(',');
}

function tryConvertMediaToContainerQuery(prelude: string): string | null {
  const afterMedia = prelude.replace(/^@media\b/i, '').trim();
  if (!afterMedia) return null;
  if (!/\b(?:min-width|max-width|width)\b/i.test(afterMedia)) return null;
  if (
    /\b(?:print|speech|prefers-|hover|any-hover|pointer|any-pointer|orientation|aspect-ratio|color|monochrome|resolution|height|min-height|max-height|scripting|update|overflow-block|overflow-inline)\b/i.test(
      afterMedia,
    )
  ) {
    return null;
  }
  const cleaned = afterMedia.replace(/\b(?:only\s+)?(?:screen|all)\s+and\s+/gi, '').trim();
  if (!cleaned) return null;
  return cleaned;
}

export function rewriteCssForShadowDom(css: string): string {
  if (!css) return '';

  let out = '';
  let i = 0;

  while (i < css.length) {
    const stmtStart = i;
    let quote: string | null = null;
    let parenDepth = 0;
    let foundBoundary: '{' | ';' | null = null;

    while (i < css.length) {
      const ch = css[i];
      if (quote) {
        if (ch === '\\') {
          i += 2;
          continue;
        }
        if (ch === quote) quote = null;
        i += 1;
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        i += 1;
        continue;
      }
      if (ch === '/' && css[i + 1] === '*') {
        const end = css.indexOf('*/', i + 2);
        i = end === -1 ? css.length : end + 2;
        continue;
      }
      if (ch === '(') parenDepth += 1;
      else if (ch === ')' && parenDepth > 0) parenDepth -= 1;
      else if (ch === ';' && parenDepth === 0) {
        foundBoundary = ';';
        break;
      } else if (ch === '{' && parenDepth === 0) {
        foundBoundary = '{';
        break;
      }
      i += 1;
    }

    if (foundBoundary === null) {
      out += css.slice(stmtStart);
      break;
    }

    if (foundBoundary === ';') {
      out += css.slice(stmtStart, i + 1);
      i += 1;
      continue;
    }

    const prelude = css.slice(stmtStart, i);
    const closeIdx = findMatchingBrace(css, i);
    const blockInner = css.slice(i + 1, closeIdx);
    const closingChar = closeIdx < css.length ? '}' : '';
    const trimmedClean = stripCssComments(prelude).trim();

    if (trimmedClean.startsWith('@')) {
      const kwMatch = trimmedClean.match(/^(?:@-[a-z-]+-|@[a-z-]+)/i);
      const kw = (kwMatch?.[0] || '').toLowerCase();

      if (VERBATIM_AT_RULES.has(kw)) {
        out += `${prelude}{${blockInner}${closingChar}`;
      } else if (kw === '@media') {
        const rewrittenInner = rewriteCssForShadowDom(blockInner);
        const containerQuery = tryConvertMediaToContainerQuery(trimmedClean);
        if (containerQuery) {
          out += `@container html-viewport ${containerQuery} {${rewrittenInner}${closingChar}`;
        } else {
          out += `${prelude}{${rewrittenInner}${closingChar}`;
        }
      } else if (RECURSIVE_AT_RULES.has(kw)) {
        out += `${prelude}{${rewriteCssForShadowDom(blockInner)}${closingChar}`;
      } else {
        out += `${prelude}{${blockInner}${closingChar}`;
      }
    } else {
      out += `${rewriteSelectorList(prelude)}{${blockInner}${closingChar}`;
    }

    i = closeIdx + 1;
  }

  return out;
}

function collectSafeAttributes(el: Element | null): Array<[string, string]> {
  if (!el) return [];
  const attrs: Array<[string, string]> = [];
  for (const attr of Array.from(el.attributes)) {
    const name = attr.name.toLowerCase();
    if (name.startsWith('on')) continue;
    if (
      name === 'class' ||
      name === 'id' ||
      name === 'style' ||
      name === 'dir' ||
      name === 'lang' ||
      name.startsWith('data-')
    ) {
      attrs.push([attr.name, attr.value]);
    }
  }
  return attrs;
}

function buildResetCss(): string {
  const varNames = new Set<string>(DOC_VIEWER_CUSTOM_VARS);
  if (typeof document !== 'undefined' && document.documentElement?.style) {
    const rootStyle = document.documentElement.style;
    for (let i = 0; i < rootStyle.length; i += 1) {
      const prop = rootStyle.item(i);
      if (prop.startsWith('--')) varNames.add(prop);
    }
  }
  const varResets = Array.from(varNames)
    .map((v) => `${v}: initial;`)
    .join(' ');

  return [
    `:host { all: initial; display: block; ${varResets} font-family: system-ui, -apple-system, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif; font-size: 16px; line-height: 1.5; color: #111827; background: #ffffff; color-scheme: light; text-align: left; }`,
    `.html-viewport-box { container-name: html-viewport; container-type: inline-size; width: 100%; min-height: 100%; display: block; }`,
    `.html-doc-html { display: block; min-height: 100%; }`,
    `.html-doc-body { display: block; min-height: 100%; margin: 8px; }`,
  ].join('\n');
}

export function buildShadowHtmlDocument(previewHtml: string): ShadowHtmlDocument {
  const doc = new DOMParser().parseFromString(previewHtml || '', 'text/html');
  const styleChunks: string[] = [];

  doc.querySelectorAll('style').forEach((styleEl) => {
    const txt = styleEl.textContent || '';
    if (txt.trim()) styleChunks.push(txt);
  });

  const bodyClone = doc.body ? (doc.body.cloneNode(true) as HTMLElement) : null;
  let bodyHtml = '';
  if (bodyClone) {
    bodyClone.querySelectorAll('style, script, noscript, iframe, object, embed, base').forEach((el) => el.remove());
    bodyHtml = bodyClone.innerHTML;
  }

  return {
    resetCss: buildResetCss(),
    stylesCss: rewriteCssForShadowDom(styleChunks.join('\n\n')),
    bodyHtml,
    htmlAttrs: collectSafeAttributes(doc.documentElement),
    bodyAttrs: collectSafeAttributes(doc.body),
  };
}
