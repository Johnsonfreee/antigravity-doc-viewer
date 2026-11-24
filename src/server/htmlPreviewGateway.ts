import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { getDefaultRootDir } from './files.js';

const MAX_MOUNTS = 32;

const MIME_MAP: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.xhtml': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.bmp': 'image/bmp',
  '.woff2': 'font/woff2',
  '.woff': 'font/woff',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.xml': 'application/xml; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
  '.wasm': 'application/wasm',
  '.mp4': 'video/mp4',
  '.webm': 'video/webm',
  '.mp3': 'audio/mpeg',
};

function normalizeForCompare(p: string): string {
  const resolved = path.resolve(p);
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

function isWithin(parent: string, child: string): boolean {
  const p = normalizeForCompare(parent);
  const c = normalizeForCompare(child);
  if (c === p) return true;
  const prefix = p.endsWith(path.sep) ? p : `${p}${path.sep}`;
  return c.startsWith(prefix);
}

function isFsRoot(dirPath: string): boolean {
  return path.dirname(dirPath) === dirPath;
}

export class HtmlPreviewGateway {
  private server: http.Server | null = null;
  private origin: string | null = null;
  private readonly tokenToRoot = new Map<string, string>();
  private readonly rootKeyToToken = new Map<string, string>();

  async start(): Promise<string | null> {
    if (this.origin) return this.origin;

    const server = http.createServer((req, res) => {
      void this.handleRequest(req, res).catch(() => {
        if (!res.headersSent) {
          res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
        }
        if (!res.writableEnded) {
          res.end('Internal server error');
        }
      });
    });

    try {
      await new Promise<void>((resolve, reject) => {
        const onError = (err: Error) => {
          server.off('listening', onListening);
          reject(err);
        };
        const onListening = () => {
          server.off('error', onError);
          resolve();
        };
        server.once('error', onError);
        server.once('listening', onListening);
        server.listen(0, '127.0.0.1');
      });

      const addr = server.address();
      if (!addr || typeof addr === 'string') {
        server.close();
        return null;
      }

      server.unref();
      this.server = server;
      this.origin = `http://127.0.0.1:${addr.port}`;
      return this.origin;
    } catch {
      try {
        server.close();
      } catch {
        // Ignore close failure after failed listen
      }
      return null;
    }
  }

  async close(): Promise<void> {
    const srv = this.server;
    this.server = null;
    this.origin = null;
    this.tokenToRoot.clear();
    this.rootKeyToToken.clear();
    if (!srv) return;

    await new Promise<void>((resolve) => {
      srv.close(() => resolve());
      srv.closeAllConnections?.();
    });
  }

  registerFile(rawFilePath: string): { liveUrl: string; staticUrl: string; origin: string } | null {
    if (!this.origin) return null;

    const filePath = path.resolve(rawFilePath);
    const fileDir = path.dirname(filePath);
    const defaultRoot = path.normalize(getDefaultRootDir());
    const homeDir = path.normalize(os.homedir());
    const homeKey = normalizeForCompare(homeDir);
    const parentDir = path.dirname(fileDir);

    let mountRoot: string;
    if (!isFsRoot(defaultRoot) && normalizeForCompare(defaultRoot) !== homeKey && isWithin(defaultRoot, filePath)) {
      mountRoot = defaultRoot;
    } else if (!isFsRoot(parentDir) && normalizeForCompare(parentDir) !== homeKey) {
      mountRoot = parentDir;
    } else {
      mountRoot = fileDir;
    }

    const mountKey = normalizeForCompare(mountRoot);
    let token = this.rootKeyToToken.get(mountKey);
    if (token) {
      this.rootKeyToToken.delete(mountKey);
      this.rootKeyToToken.set(mountKey, token);
      this.tokenToRoot.delete(token);
      this.tokenToRoot.set(token, mountRoot);
    } else {
      if (this.tokenToRoot.size >= MAX_MOUNTS) {
        const oldestToken = this.tokenToRoot.keys().next().value;
        if (oldestToken) {
          const oldestRoot = this.tokenToRoot.get(oldestToken);
          this.tokenToRoot.delete(oldestToken);
          if (oldestRoot) this.rootKeyToToken.delete(normalizeForCompare(oldestRoot));
        }
      }
      token = crypto.randomBytes(16).toString('hex');
      this.tokenToRoot.set(token, mountRoot);
      this.rootKeyToToken.set(mountKey, token);
    }

    const relUrlPath = path
      .relative(mountRoot, filePath)
      .split(path.sep)
      .filter(Boolean)
      .map(encodeURIComponent)
      .join('/');

    return {
      origin: this.origin,
      liveUrl: `${this.origin}/p/${token}/${relUrlPath}?mode=live`,
      staticUrl: `${this.origin}/p/${token}/${relUrlPath}?mode=static`,
    };
  }

  private async handleRequest(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const method = (req.method || 'GET').toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
      res.writeHead(405, {
        Allow: 'GET, HEAD',
        'Content-Type': 'text/plain; charset=utf-8',
        'Cache-Control': 'no-store',
        'X-Content-Type-Options': 'nosniff',
      });
      res.end('Method Not Allowed');
      return;
    }

    const origin = this.origin || 'http://127.0.0.1';
    let parsedUrl: URL;
    try {
      parsedUrl = new URL(req.url || '/', origin);
    } catch {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Bad Request');
      return;
    }

    const m = /^\/p\/([a-f0-9]{32})\/(.+)$/.exec(parsedUrl.pathname);
    if (!m) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }

    const token = m[1];
    const mountRoot = this.tokenToRoot.get(token);
    if (!mountRoot) {
      res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Forbidden');
      return;
    }

    const segments: string[] = [];
    for (const rawSeg of m[2].split('/')) {
      if (!rawSeg) {
        res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Invalid Path');
        return;
      }
      let decoded: string;
      try {
        decoded = decodeURIComponent(rawSeg);
      } catch {
        res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Bad URI Encoding');
        return;
      }
      if (!decoded || decoded.includes('\0') || decoded.includes('/') || decoded.includes('\\')) {
        res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Forbidden');
        return;
      }
      segments.push(decoded);
    }

    const candidate = path.resolve(mountRoot, ...segments);
    if (!isWithin(mountRoot, candidate)) {
      res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Forbidden');
      return;
    }

    let stat;
    try {
      stat = await fs.stat(candidate);
    } catch {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }
    if (!stat.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }

    try {
      const [realMount, realTarget] = await Promise.all([fs.realpath(mountRoot), fs.realpath(candidate)]);
      if (!isWithin(realMount, realTarget)) {
        res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('Forbidden');
        return;
      }
    } catch {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }

    const ext = path.extname(candidate).toLowerCase();
    const contentType = MIME_MAP[ext] || 'application/octet-stream';
    const headers: Record<string, string | number> = {
      'Content-Type': contentType,
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
    };

    if (ext === '.html' || ext === '.htm' || ext === '.xhtml') {
      if (parsedUrl.searchParams.get('mode') === 'static') {
        headers['Content-Security-Policy'] =
          "default-src 'self'; script-src 'none'; style-src 'self' 'unsafe-inline' data:; img-src 'self' data: blob:; font-src 'self' data: blob:; connect-src 'none'; frame-src 'none'; object-src 'none';";
      } else {
        headers['Content-Security-Policy'] =
          "default-src 'self' https: http: data: blob: 'unsafe-inline' 'unsafe-eval'; frame-ancestors *;";
      }
    }

    const body = method === 'HEAD' ? null : await fs.readFile(candidate);
    if (body) {
      headers['Content-Length'] = body.byteLength;
    }

    res.writeHead(200, headers);
    res.end(body ?? undefined);
  }
}
