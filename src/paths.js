import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';

/**
 * Canonical filesystem locations for the Teams Browser MCP.
 * Matches the C# `Paths` type so state.json lives in the same place
 * across the C#→Node rewrite (state files carry over cleanly).
 */
const root = path.join(
  process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'),
  'TeamsBrowserMcp',
);

fs.mkdirSync(root, { recursive: true });

const logsDir = path.join(root, 'logs');
fs.mkdirSync(logsDir, { recursive: true });

export const Paths = {
  root,
  stateFile: path.join(root, 'state.json'),
  logsDir,
  logFile: path.join(logsDir, 'mcp.log'),
};
