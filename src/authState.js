import fs from 'node:fs';
import { Paths } from './paths.js';
import { log } from './logger.js';

/**
 * state.json is a Playwright storageState JSON document written and read
 * by Playwright itself via `context.storageState({ path })` and the
 * `storageState:` newContext option. This module is just a thin wrapper
 * for existence/deletion — validity is decided by actually loading the
 * page and checking the logged-in DOM markers.
 */
export const AuthState = {
  exists() {
    return fs.existsSync(Paths.stateFile);
  },

  delete() {
    try {
      fs.unlinkSync(Paths.stateFile);
      log.info('authState', 'deleted', { path: Paths.stateFile });
    } catch (err) {
      if (err.code !== 'ENOENT') throw err;
    }
  },
};
