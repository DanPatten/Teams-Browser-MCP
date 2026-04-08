import { session } from '../teamsSession.js';
import { log } from '../logger.js';

export const authenticateTool = {
  name: 'authenticate',
  description:
    'Force a re-authentication cycle. Deletes any saved Teams session ' +
    'state, relaunches the Playwright Chrome window on the Teams sign-in ' +
    'page, and blocks until the user signs in inside that window. Call ' +
    'this once at the start of a session, or after another tool returns ' +
    'an auth-expired error. Blocks until login completes (up to 5 minutes).',
  inputSchema: { type: 'object', properties: {}, additionalProperties: false },

  async handler() {
    return log.span('tool.authenticate', 'tool', async () => {
      await session.forceReauthenticate();
      return { content: [{ type: 'text', text: 'ok: Teams session is ready' }] };
    });
  },
};
