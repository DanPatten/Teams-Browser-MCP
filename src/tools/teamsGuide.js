import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { log } from '../logger.js';

const __filename = fileURLToPath(import.meta.url);
const GUIDE_PATH = path.resolve(path.dirname(__filename), '..', '..', 'TEAMS_GUIDE.md');

export const teamsGuideTool = {
  name: 'teams_guide',
  description:
    'Returns TEAMS_GUIDE.md — a navigation guide for driving Microsoft ' +
    'Teams via this MCP\'s primitive tools (teams_query, teams_click, ' +
    'teams_type, etc.). Read this once at the start of a session before ' +
    'attempting to use Teams.',
  inputSchema: { type: 'object', properties: {}, additionalProperties: false },

  async handler() {
    return log.span('tool.teams_guide', 'tool', async () => {
      const text = fs.readFileSync(GUIDE_PATH, 'utf8');
      log.info('tool.teams_guide', 'served', { bytes: text.length });
      return { content: [{ type: 'text', text }] };
    });
  },
};
