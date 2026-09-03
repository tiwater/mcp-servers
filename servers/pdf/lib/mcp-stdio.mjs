import process from 'node:process';

const JSONRPC_VERSION = '2.0';
const SUPPORTED_PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05', '2024-10-07'];
const FALLBACK_PROTOCOL_VERSION = '2025-06-18';

function writeMessage(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function toError(code, message, data) {
  return { code, message, ...(data === undefined ? {} : { data }) };
}

function normalizeToolCallError(error) {
  if (Number.isInteger(error?.code) && error.message) {
    return toError(error.code, error.message, error.data);
  }
  return toError(-32603, error instanceof Error ? error.message : String(error || 'Unknown error'));
}

export class McpStdioServer {
  constructor({ name, version, instructions, tools, callTool, logger = console.error }) {
    this.serverInfo = { name, version };
    this.instructions = instructions;
    this.tools = tools;
    this.callTool = callTool;
    this.logger = logger;
    this.lineBuffer = '';
    this.binaryBuffer = Buffer.alloc(0);
  }

  start() {
    process.stdin.on('data', chunk => this.#onData(chunk));
    process.stdin.on('end', () => process.exit(0));
  }

  #onData(chunk) {
    const text = chunk.toString('utf8');
    if (this.binaryBuffer.length > 0 || text.includes('Content-Length:')) {
      this.binaryBuffer = Buffer.concat([this.binaryBuffer, chunk]);
      this.#drainContentLengthBuffer();
      return;
    }
    this.lineBuffer += text;
    while (true) {
      const newlineIndex = this.lineBuffer.indexOf('\n');
      if (newlineIndex === -1) return;
      const line = this.lineBuffer.slice(0, newlineIndex).replace(/\r$/u, '').trim();
      this.lineBuffer = this.lineBuffer.slice(newlineIndex + 1);
      if (line) this.#parseAndHandle(line, null);
    }
  }

  #drainContentLengthBuffer() {
    while (true) {
      const headerEnd = this.binaryBuffer.indexOf('\r\n\r\n');
      if (headerEnd === -1) return;
      const headerText = this.binaryBuffer.subarray(0, headerEnd).toString('utf8');
      const lengthMatch = headerText.match(/Content-Length:\s*(\d+)/iu);
      if (!lengthMatch) {
        this.logger('Missing Content-Length header');
        this.binaryBuffer = Buffer.alloc(0);
        return;
      }
      const messageStart = headerEnd + 4;
      const messageEnd = messageStart + Number(lengthMatch[1]);
      if (this.binaryBuffer.length < messageEnd) return;
      const body = this.binaryBuffer.subarray(messageStart, messageEnd).toString('utf8');
      this.binaryBuffer = this.binaryBuffer.subarray(messageEnd);
      this.#parseAndHandle(body, null);
    }
  }

  #parseAndHandle(body, idHint) {
    try {
      void this.#handleMessage(JSON.parse(body));
    } catch (error) {
      writeMessage({ jsonrpc: JSONRPC_VERSION, id: idHint, error: toError(-32700, 'Parse error', String(error)) });
    }
  }

  async #handleMessage(message) {
    if (!message || message.jsonrpc !== JSONRPC_VERSION || typeof message.method !== 'string') {
      if ('id' in (message || {})) {
        writeMessage({ jsonrpc: JSONRPC_VERSION, id: message.id ?? null, error: toError(-32600, 'Invalid Request') });
      }
      return;
    }
    const { id, method, params = {} } = message;
    const isNotification = id === undefined;
    try {
      if (method === 'initialize') {
        const protocolVersion = SUPPORTED_PROTOCOL_VERSIONS.includes(params.protocolVersion)
          ? params.protocolVersion : FALLBACK_PROTOCOL_VERSION;
        if (!isNotification) writeMessage({
          jsonrpc: JSONRPC_VERSION,
          id,
          result: {
            protocolVersion,
            capabilities: { tools: {} },
            serverInfo: this.serverInfo,
            ...(this.instructions ? { instructions: this.instructions } : {}),
          },
        });
      } else if (method === 'notifications/initialized') {
        // No response for notifications.
      } else if (method === 'ping') {
        if (!isNotification) writeMessage({ jsonrpc: JSONRPC_VERSION, id, result: {} });
      } else if (method === 'tools/list') {
        if (!isNotification) writeMessage({ jsonrpc: JSONRPC_VERSION, id, result: { tools: this.tools } });
      } else if (method === 'tools/call') {
        const name = params?.name;
        if (typeof name !== 'string' || !name) {
          throw Object.assign(new Error('Invalid params: missing tool name'), { code: -32602 });
        }
        const result = await this.callTool(name, params.arguments ?? {});
        if (!isNotification) writeMessage({ jsonrpc: JSONRPC_VERSION, id, result });
      } else if (!isNotification) {
        writeMessage({ jsonrpc: JSONRPC_VERSION, id, error: toError(-32601, `Method not found: ${method}`) });
      }
    } catch (error) {
      if (!isNotification) writeMessage({ jsonrpc: JSONRPC_VERSION, id, error: normalizeToolCallError(error) });
    }
  }
}
