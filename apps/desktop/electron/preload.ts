/** Exposes the typed DeckForge API to the sandboxed renderer. Nothing else crosses over. */
import { contextBridge, ipcRenderer } from 'electron';
import type { DeckforgeApi } from './api.ts';

const call =
  <R>(channel: string) =>
  (...args: unknown[]): Promise<R> =>
    ipcRenderer.invoke(channel, ...args).catch((e: Error) => {
      // Electron prefixes errors with "Error invoking remote method '<channel>': Error: ".
      throw new Error(e.message.replace(/^Error invoking remote method '[^']+': (Error: )?/, ''));
    });

const api: DeckforgeApi = {
  pickTemplate: call('pickTemplate'),
  pickData: call('pickData'),
  pickFolder: call('pickFolder'),
  inspect: call('inspect'),
  createDataFile: call('createDataFile'),
  check: call('check'),
  generate: call('generate'),
  openPath: call('openPath'),
  showInFolder: call('showInFolder'),
  previewTool: call('previewTool'),
  preview: call('preview'),
  recipes: {
    list: call('recipes.list'),
    save: call('recipes.save'),
    remove: call('recipes.remove'),
  },
  license: {
    status: call('license.status'),
    activate: call('license.activate'),
    deactivate: call('license.deactivate'),
    requestCode: call('license.requestCode'),
  },
  updates: {
    status: call('updates.status'),
    check: call('updates.check'),
    setAuto: call('updates.setAuto'),
    install: call('updates.install'),
  },
};

contextBridge.exposeInMainWorld('deckforge', api);
