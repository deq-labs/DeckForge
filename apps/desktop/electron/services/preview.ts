/**
 * Rendered preview: converts a deck to PDF with PowerPoint (exact) or LibreOffice (close), if
 * one is installed. Nothing is uploaded anywhere; both run locally.
 */
import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { basename, join } from 'node:path';
import type { PreviewTool } from '../api.ts';

export interface ToolPaths {
  powerpoint?: string;
  libreoffice?: string;
}

/** Where the tools usually live; pure so it can be tested per platform. */
export function candidatePaths(platform: NodeJS.Platform, env: Record<string, string | undefined>): { powerpoint: string[]; libreoffice: string[] } {
  if (platform === 'win32') {
    const pf = [env.ProgramFiles ?? 'C:\\Program Files', env['ProgramFiles(x86)'] ?? 'C:\\Program Files (x86)'];
    return {
      powerpoint: pf.flatMap((p) => [`${p}\\Microsoft Office\\root\\Office16\\POWERPNT.EXE`, `${p}\\Microsoft Office\\Office16\\POWERPNT.EXE`]),
      libreoffice: pf.map((p) => `${p}\\LibreOffice\\program\\soffice.exe`),
    };
  }
  if (platform === 'darwin') {
    return {
      powerpoint: ['/Applications/Microsoft PowerPoint.app'],
      libreoffice: ['/Applications/LibreOffice.app/Contents/MacOS/soffice'],
    };
  }
  return { powerpoint: [], libreoffice: ['/usr/bin/soffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice', '/opt/libreoffice/program/soffice'] };
}

export function detectTools(platform = process.platform, env: Record<string, string | undefined> = process.env, exists: (p: string) => boolean = existsSync): ToolPaths {
  const c = candidatePaths(platform, env);
  const powerpoint = c.powerpoint.find(exists);
  const libreoffice = c.libreoffice.find(exists);
  return { ...(powerpoint ? { powerpoint } : {}), ...(libreoffice ? { libreoffice } : {}) };
}

/** PowerPoint renders exactly what users will see, so it is preferred when present. */
export function chooseTool(tools: ToolPaths): PreviewTool | null {
  if (tools.powerpoint) return 'powerpoint';
  if (tools.libreoffice) return 'libreoffice';
  return null;
}

function exec(file: string, args: string[], timeoutMs: number): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(file, args, { timeout: timeoutMs, windowsHide: true }, (err, _stdout, stderr) => {
      if (err) reject(new Error(`${basename(file)} could not render the preview: ${stderr?.toString().trim() || err.message}`));
      else resolve();
    });
  });
}

/** Single-quoted PowerShell string. */
const psq = (s: string): string => `'${s.replace(/'/g, "''")}'`;
/** AppleScript string literal. */
const asq = (s: string): string => `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;

export async function renderPdf(deckPath: string, outDir: string, tools: ToolPaths = detectTools()): Promise<string> {
  await mkdir(outDir, { recursive: true });
  const pdf = join(outDir, `${basename(deckPath).replace(/\.pptx$/i, '')}.pdf`);
  const tool = chooseTool(tools);
  if (tool === 'powerpoint' && process.platform === 'win32') {
    // Open read-only without a window, save as PDF (32 = ppSaveAsPDF), close.
    const script =
      `$a = New-Object -ComObject PowerPoint.Application; ` +
      `$p = $a.Presentations.Open(${psq(deckPath)}, -1, 0, 0); ` +
      `try { $p.SaveAs(${psq(pdf)}, 32) } finally { $p.Close() }`;
    await exec('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], 120_000);
  } else if (tool === 'powerpoint' && process.platform === 'darwin') {
    const script = [
      'tell application "Microsoft PowerPoint"',
      `  set d to open (POSIX file ${asq(deckPath)})`,
      `  save d in (POSIX file ${asq(pdf)}) as save as PDF`,
      '  close d saving no',
      'end tell',
    ].join('\n');
    await exec('osascript', ['-e', script], 120_000);
  } else if (tool === 'libreoffice' && tools.libreoffice) {
    await exec(tools.libreoffice, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', outDir, deckPath], 120_000);
  } else {
    throw new Error('Preview needs PowerPoint or LibreOffice installed on this computer. The decks themselves are fine without it.');
  }
  if (!existsSync(pdf)) throw new Error('The preview could not be rendered. Open the deck directly instead.');
  return pdf;
}
