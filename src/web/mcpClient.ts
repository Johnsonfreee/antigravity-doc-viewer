import {
  App,
  applyDocumentTheme,
  applyHostStyleVariables,
  type McpUiHostCapabilities,
  type McpUiHostContext,
} from '@modelcontextprotocol/ext-apps';
import { z } from 'zod/v4';

export interface ToolResultHandlers {
  onToolInput?: (args: Record<string, unknown>) => void;
  onToolResult?: (structured: Record<string, unknown> | undefined, isError?: boolean, errorText?: string) => void;
  onToolCancelled?: (reason?: string) => void;
}

const MAX_CONCURRENT_CALLS = 3;

export class ViewerMcpClient {
  private readonly app = new App({ name: 'DocViewer', version: '1.0.0' }, {});
  private inFlight = 0;
  private readonly waitQueue: Array<() => void> = [];
  private connected = false;

  async connect(handlers: ToolResultHandlers): Promise<void> {
    this.app.ontoolinput = (params) => {
      handlers.onToolInput?.((params.arguments ?? {}) as Record<string, unknown>);
    };

    this.app.ontoolresult = (result) => {
      const textMsg = result.content
        ?.filter((c): c is { type: 'text'; text: string } => c.type === 'text')
        .map((c) => c.text)
        .join('\n');
      handlers.onToolResult?.(
        result.structuredContent as Record<string, unknown> | undefined,
        result.isError,
        textMsg,
      );
    };

    this.app.ontoolcancelled = (params) => {
      handlers.onToolCancelled?.(params.reason);
    };

    this.app.onhostcontextchanged = (ctx) => {
      this.applyThemeContext(ctx);
    };

    await this.app.connect();
    this.connected = true;
    const ctx = this.app.getHostContext();
    if (ctx) {
      this.applyThemeContext(ctx);
    }
  }

  private applyThemeContext(ctx: McpUiHostContext): void {
    if (ctx.theme) {
      applyDocumentTheme(ctx.theme);
    }
    if (ctx.styles?.variables) {
      applyHostStyleVariables(ctx.styles.variables);
    }
  }

  getCapabilities(): McpUiHostCapabilities | undefined {
    return this.app.getHostCapabilities();
  }

  canMessage(): boolean {
    return Boolean(this.connected && this.app.getHostCapabilities()?.message);
  }

  canUpdateContext(): boolean {
    return Boolean(this.connected && this.app.getHostCapabilities()?.updateModelContext);
  }

  private async acquireSlot(): Promise<void> {
    if (this.inFlight < MAX_CONCURRENT_CALLS) {
      this.inFlight += 1;
      return;
    }
    await new Promise<void>((resolve) => this.waitQueue.push(resolve));
    this.inFlight += 1;
  }

  private releaseSlot(): void {
    this.inFlight = Math.max(0, this.inFlight - 1);
    const next = this.waitQueue.shift();
    if (next) next();
  }

  async callTool<T = Record<string, unknown>>(
    name: string,
    args: Record<string, unknown> = {},
  ): Promise<T> {
    await this.acquireSlot();
    try {
      const result = await this.app.callServerTool({ name, arguments: args });
      const textBlock = result.content?.find(
        (c): c is { type: 'text'; text: string } => c.type === 'text',
      )?.text;

      if (result.isError) {
        throw new Error(textBlock || `Tool ${name} failed`);
      }
      if (result.structuredContent !== undefined) {
        return result.structuredContent as T;
      }
      if (textBlock) {
        return JSON.parse(textBlock) as T;
      }
      return {} as T;
    } finally {
      this.releaseSlot();
    }
  }

  async attachContext(text: string): Promise<void> {
    if (!this.canUpdateContext()) {
      throw new Error('当前视图不支持附加对话上下文');
    }
    await this.app.updateModelContext({
      content: [{ type: 'text', text }],
    });
  }

  async sendMessage(text: string, contextText?: string): Promise<void> {
    if (!this.canMessage()) {
      throw new Error('当前视图（如全局侧边栏）不支持直接发送对话消息');
    }
    if (contextText && this.canUpdateContext()) {
      await this.app.updateModelContext({
        content: [{ type: 'text', text: contextText }],
      });
    }
    const res = await this.app.sendMessage({
      role: 'user',
      content: [{ type: 'text', text }],
    });
    if (res.isError) {
      throw new Error('宿主拒绝了消息发送');
    }
  }

  async openInNativeFileView(filePath: string): Promise<boolean> {
    if (!this.connected) return false;
    const normalized = filePath.replace(/\\/g, '/');
    const uri = normalized.startsWith('/')
      ? `file://${encodeURI(normalized)}`
      : `file:///${encodeURI(normalized)}`;
    try {
      const req = {
        method: 'com.google.antigravity/toggle-aux-pane',
        params: { tabType: 'fileView', uri },
      } as unknown as Parameters<App['request']>[0];
      const res = await this.app.request(
        req,
        z.object({ isError: z.boolean().optional() }).passthrough(),
      );
      return !(res as { isError?: boolean }).isError;
    } catch {
      return false;
    }
  }
}
