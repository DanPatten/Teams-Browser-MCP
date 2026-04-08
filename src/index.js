#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { log } from './logger.js';
import { authenticateTool } from './tools/authenticate.js';
import { teamsGuideTool } from './tools/teamsGuide.js';
import {
  teamsNavigateTool,
  teamsQueryTool,
  teamsClickTool,
  teamsTypeTool,
  teamsPressKeyTool,
  teamsWaitForTool,
  teamsEvaluateTool,
} from './tools/primitives.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgPath = path.join(__dirname, '..', 'package.json');
const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));

// Boot line — prints immediately so the user can always confirm which
// source the running MCP is on. Includes source mtime so "did my edit
// land?" is answered by a quick log tail. This is the explicit fix for
// the old C# problem of silently running a stale DLL.
const selfMtime = fs.statSync(__filename).mtime.toISOString();
log.info('server', 'boot', {
  pid: process.pid,
  node: process.version,
  cwd: process.cwd(),
  entry: __filename,
  entryMtime: selfMtime,
  version: pkg.version,
});

const tools = [
  authenticateTool,
  teamsGuideTool,
  teamsNavigateTool,
  teamsQueryTool,
  teamsClickTool,
  teamsTypeTool,
  teamsPressKeyTool,
  teamsWaitForTool,
  teamsEvaluateTool,
];
const toolMap = new Map(tools.map((t) => [t.name, t]));

const server = new Server(
  { name: 'teams-browser-mcp', version: pkg.version },
  { capabilities: { tools: {} } },
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  log.info('server', 'list-tools', { count: tools.length });
  return {
    tools: tools.map((t) => ({
      name: t.name,
      description: t.description,
      inputSchema: t.inputSchema,
    })),
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  const tool = toolMap.get(name);
  if (!tool) {
    log.warn('server', 'unknown-tool', { name });
    return {
      content: [{ type: 'text', text: `Unknown tool: ${name}` }],
      isError: true,
    };
  }

  log.info('server', 'tool:call', { name, args: summarizeArgs(args) });
  try {
    const result = await tool.handler(args || {});
    log.info('server', 'tool:call:ok', { name });
    return result;
  } catch (err) {
    log.error('server', 'tool:call:failed', err, { name });
    return {
      content: [{ type: 'text', text: err?.message || String(err) }],
      isError: true,
    };
  }
});

function summarizeArgs(args) {
  if (!args || typeof args !== 'object') return '{}';
  const out = {};
  for (const [k, v] of Object.entries(args)) {
    if (typeof v === 'string' && v.length > 120) out[k] = v.slice(0, 120) + '…';
    else out[k] = v;
  }
  return JSON.stringify(out);
}

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log.info('server', 'stdio-connected');
}

process.on('uncaughtException', (err) => {
  log.error('server', 'uncaughtException', err);
});
process.on('unhandledRejection', (err) => {
  log.error('server', 'unhandledRejection', err);
});

main().catch((err) => {
  log.error('server', 'main:failed', err);
  process.exit(1);
});
