/** Electron main process: window, IPC handlers, file dialogs. All file work happens here. */
import { BrowserWindow, Menu, app, dialog, ipcMain, shell } from 'electron';
import type { IpcMainInvokeEvent } from 'electron';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import type { JobSpec, Recipe } from './api.ts';
import { WATERMARK, checkJob, generateJob, inspectTemplate, writeDataFile } from './services/jobs.ts';
import { RecipeStore } from './services/recipes.ts';
import { chooseTool, detectTools, renderPdf } from './services/preview.ts';
import { LicenseService } from './services/license.ts';
import type { LicenseConfig } from './services/license.ts';
import { UpdateService } from './services/updates.ts';
import type { Updater } from './services/updates.ts';
import { LICENSE_SERVER_URL, PRODUCTION_PUBLIC_KEY } from '@deckforge/licensing';
import { autoUpdater } from 'electron-updater';

/** Set at build time by scripts/build-electron.mjs (YYYY-MM-DD); empty in development. */
declare const __DECKFORGE_RELEASE_DATE__: string;

const devServer = process.env.VITE_DEV_SERVER_URL;
let win: BrowserWindow | null = null;
let recipes: RecipeStore;
let license: LicenseService;
let updates: UpdateService;

/** Production values, with overrides for development and tests only (never in installed builds). */
function licenseConfig(): LicenseConfig {
  const dev = !app.isPackaged;
  const release = typeof __DECKFORGE_RELEASE_DATE__ === 'string' && __DECKFORGE_RELEASE_DATE__ ? __DECKFORGE_RELEASE_DATE__ : undefined;
  return {
    publicKey: (dev && process.env.DECKFORGE_LICENSE_PUBLIC_KEY) || PRODUCTION_PUBLIC_KEY,
    releaseDate: (dev && process.env.DECKFORGE_RELEASE_DATE) || (app.isPackaged ? release : undefined),
    serverUrl: (dev && process.env.DECKFORGE_LICENSE_SERVER_URL) || LICENSE_SERVER_URL,
  };
}

/** Paths the UI may ask the OS to open: files the user picked or DeckForge wrote. */
const openable = new Set<string>();
const allow = (p: string): string => {
  openable.add(resolve(p).toLowerCase());
  return p;
};
const isAllowed = (p: string): boolean => {
  const r = resolve(p).toLowerCase();
  return [...openable].some((a) => r === a || r.startsWith(`${a}\\`) || r.startsWith(`${a}/`));
};

function createWindow(): void {
  win = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 600,
    title: 'DeckForge',
    // Installed builds take the icon from the executable; this covers development runs.
    ...(app.isPackaged ? {} : { icon: join(__dirname, '..', 'build', 'icon.png') }),
    backgroundColor: '#f6f7f9',
    webPreferences: {
      preload: join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      // Smoke mode renders offscreen so screenshots work without a visible, focused window.
      ...(process.env.DECKFORGE_SMOKE ? { offscreen: true } : {}),
    },
    ...(process.env.DECKFORGE_SMOKE ? { show: false } : {}),
  });
  // Links never navigate the app window; external URLs open in the browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https:\/\//.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  if (devServer) void win.loadURL(devServer);
  else void win.loadFile(join(__dirname, '..', 'dist', 'index.html'));
}

function handle<A extends unknown[], R>(channel: string, fn: (...args: A) => Promise<R> | R): void {
  ipcMain.handle(channel, async (_e: IpcMainInvokeEvent, ...args: unknown[]) => fn(...(args as A)));
}

function assertSpec(spec: JobSpec): JobSpec {
  if (!spec?.templatePath || !spec.dataPath || !spec.outDir) throw new Error('Choose a template, a data workbook and an output folder first.');
  return spec;
}

function registerIpc(): void {
  handle('pickTemplate', async () => {
    const r = await dialog.showOpenDialog(win!, { title: 'Choose a PowerPoint template', filters: [{ name: 'PowerPoint', extensions: ['pptx'] }], properties: ['openFile'] });
    return r.canceled || !r.filePaths[0] ? null : allow(r.filePaths[0]);
  });
  handle('pickData', async () => {
    const r = await dialog.showOpenDialog(win!, { title: 'Choose a data workbook', filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }], properties: ['openFile'] });
    return r.canceled || !r.filePaths[0] ? null : allow(r.filePaths[0]);
  });
  handle('pickFolder', async (defaultPath?: string) => {
    const r = await dialog.showOpenDialog(win!, { title: 'Choose where to save the decks', ...(defaultPath ? { defaultPath } : {}), properties: ['openDirectory', 'createDirectory'] });
    return r.canceled || !r.filePaths[0] ? null : allow(r.filePaths[0]);
  });
  handle('inspect', async (templatePath: string) => {
    allow(templatePath);
    return inspectTemplate(templatePath);
  });
  handle('createDataFile', async (templatePath: string) => {
    const stem = basename(templatePath, extname(templatePath));
    const r = await dialog.showSaveDialog(win!, {
      title: 'Save the data workbook',
      defaultPath: join(dirname(templatePath), `${stem}.data.xlsx`),
      filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
    });
    if (r.canceled || !r.filePath) return null;
    await writeDataFile(templatePath, r.filePath);
    allow(r.filePath);
    void shell.openPath(r.filePath);
    return r.filePath;
  });
  handle('check', async (spec: JobSpec) => checkJob(assertSpec(spec)));
  handle('generate', async (spec: JobSpec) => {
    const lic = await license.status();
    const result = await generateJob(assertSpec(spec), lic.mode === 'licensed' ? undefined : WATERMARK);
    allow(spec.outDir);
    result.files.forEach(allow);
    return result;
  });
  handle('openPath', async (p: string) => {
    if (!isAllowed(p) || !existsSync(p)) throw new Error('That file is no longer available.');
    const err = await shell.openPath(p);
    if (err) throw new Error(err);
  });
  handle('showInFolder', async (p: string) => {
    if (!isAllowed(p)) throw new Error('That file is no longer available.');
    shell.showItemInFolder(p);
  });
  handle('previewTool', () => chooseTool(detectTools()));
  handle('preview', async (deckPath: string) => {
    if (!isAllowed(deckPath)) throw new Error('That deck is no longer available.');
    const pdf = await renderPdf(deckPath, join(tmpdir(), 'deckforge-preview'));
    const viewer = new BrowserWindow({
      width: 1100,
      height: 760,
      title: `Preview — ${basename(deckPath)}`,
      webPreferences: { plugins: true, contextIsolation: true, nodeIntegration: false, sandbox: true },
    });
    viewer.setMenuBarVisibility(false);
    await viewer.loadFile(pdf); // Chromium's built-in PDF viewer
  });
  handle('recipes.list', async () => {
    const list = await recipes.list();
    for (const r of list) [r.spec.templatePath, r.spec.dataPath, r.spec.outDir].forEach(allow);
    return list;
  });
  handle('recipes.save', async (r: Recipe) => {
    [r.spec.templatePath, r.spec.dataPath, r.spec.outDir].forEach(allow);
    return recipes.save(r);
  });
  handle('recipes.remove', async (id: string) => recipes.remove(id));
  handle('license.status', () => license.status());
  handle('license.activate', (text: string, opts?: { online?: boolean }) => license.activate(text, opts ?? {}));
  handle('license.deactivate', () => license.deactivate());
  handle('license.requestCode', () => license.requestCode());
  handle('updates.status', () => updates.status());
  handle('updates.check', () => updates.check());
  handle('updates.setAuto', (auto: boolean) => updates.setAuto(auto));
  handle('updates.install', () => updates.install());
}

/**
 * Smoke mode (tests/CI only): DECKFORGE_SMOKE=<json> drives the real renderer + IPC end to end,
 * writes a screenshot and a JSON report, then quits. Never set in normal use.
 */
async function smoke(cfg: { template: string; data: string; outDir: string; screenshot: string; report: string; licenseKey?: string }): Promise<void> {
  const { writeFile } = await import('node:fs/promises');
  const w = win!;
  await new Promise<void>((r) => (w.webContents.isLoading() ? w.webContents.once('did-finish-load', () => r()) : r()));
  [cfg.template, cfg.data, cfg.outDir].forEach(allow);
  const spec: JobSpec = { templatePath: cfg.template, dataPath: cfg.data, outDir: cfg.outDir, options: { locale: 'en-US', currency: 'USD', missing: 'error' } };
  const js = `(async () => {
    const api = window.deckforge;
    const activated = ${cfg.licenseKey ? `await api.license.activate(${JSON.stringify(cfg.licenseKey)})` : 'null'};
    const info = await api.inspect(${JSON.stringify(cfg.template)});
    const checked = await api.check(${JSON.stringify(spec)});
    const generated = await api.generate(${JSON.stringify(spec)});
    const lic = await api.license.status();
    return { hasApi: !!api, activatedMode: activated && activated.mode, slots: info.slots.length, checkOk: checked.ok, files: generated.files, watermarked: generated.watermarked, license: lic.mode };
  })()`;
  let report: unknown;
  try {
    report = await w.webContents.executeJavaScript(js);
  } catch (e) {
    report = { error: (e as Error).message };
  }
  await writeFile(cfg.report, JSON.stringify(report, null, 2));
  // Reopen the UI with the job loaded through the same path as a recipe, then capture it.
  await w.loadFile(join(__dirname, '..', 'dist', 'index.html'), { hash: `job=${encodeURIComponent(JSON.stringify(spec))}` });
  // Same document + new hash is an in-page navigation; reload so the app starts fresh with the job.
  await new Promise<void>((r) => {
    w.webContents.once('did-finish-load', () => r());
    w.webContents.reload();
  });
  w.setContentSize(1200, 1500);
  await new Promise((r) => setTimeout(r, 4000));
  const shot = await Promise.race([w.webContents.capturePage(), new Promise<null>((r) => setTimeout(() => r(null), 10_000))]);
  if (shot) await writeFile(cfg.screenshot, shot.toPNG());
  app.quit();
}

app.whenReady().then(() => {
  const smokeCfg = process.env.DECKFORGE_SMOKE;
  const dataDir = smokeCfg ? join(tmpdir(), `deckforge-smoke-${process.pid}`) : app.getPath('userData');
  recipes = new RecipeStore(dataDir);
  license = new LicenseService(dataDir, licenseConfig());
  updates = new UpdateService(dataDir, app.isPackaged ? (autoUpdater as unknown as Updater) : undefined, async () => {
    const st = await license.state();
    return st.mode === 'licensed' ? st.license?.updatesUntil : undefined;
  });
  void updates.init();
  if (!devServer) Menu.setApplicationMenu(null);
  registerIpc();
  createWindow();
  if (smokeCfg) void smoke(JSON.parse(smokeCfg));
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
