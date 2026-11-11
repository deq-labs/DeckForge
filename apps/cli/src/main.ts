#!/usr/bin/env node
import { Command, Option } from 'commander';
import { batch, check, fill, initData, license, slots } from './commands.ts';
import { UserError } from './files.ts';
import { consoleIo } from './report.ts';

const program = new Command()
  .name('deckforge')
  .description('Fill your own PowerPoint template with data from an Excel workbook.')
  .version('0.0.0')
  .showHelpAfterError();

program
  .command('slots')
  .description('List the tags, tables and charts a template can be filled with, and any problems in it.')
  .argument('<template>', 'tagged PowerPoint template (.pptx)')
  .option('--json', 'print the slot manifest as JSON')
  .action((template: string, opts) => run(() => slots(template, opts, consoleIo)));

program
  .command('init-data')
  .description('Create a data workbook laid out for a template, ready to fill in.')
  .argument('<template>', 'tagged PowerPoint template (.pptx)')
  .option('-o, --output <file>', 'workbook to create (default: <template>.data.xlsx)')
  .option('--no-sample', 'leave tables and charts empty instead of copying the template\'s current data')
  .option('-f, --force', 'replace the workbook if it already exists')
  .action((template: string, opts) => run(() => initData(template, opts, consoleIo)));

program
  .command('fill')
  .description('Fill a template with a data workbook and write the deck.')
  .argument('<template>', 'tagged PowerPoint template (.pptx)')
  .argument('<data>', 'data workbook (.xlsx)')
  .option('-o, --output <file>', 'deck to write (default: <template>.filled.pptx); replaced if it exists')
  .option('--locale <code>', 'number and date style, e.g. en-US, en-GB, de-DE', 'en-US')
  .option('--currency <code>', 'currency for the "currency" format, e.g. USD, EUR', 'USD')
  .addOption(new Option('--missing <mode>', 'what to do with tags that have no data').choices(['error', 'keep', 'blank']).default('error'))
  .action((template: string, data: string, opts) => run(() => fill(template, data, opts, consoleIo)));

program
  .command('batch')
  .description('Write one deck per value of the "batch" column in the workbook.')
  .argument('<template>', 'tagged PowerPoint template (.pptx)')
  .argument('<data>', 'data workbook (.xlsx) with a "batch" column on one or more sheets')
  .option('-d, --out-dir <folder>', 'folder for the decks (default: <template>-decks next to the template)')
  .option('--name <pattern>', 'file name pattern; {template} and {batch} are replaced', '{template}-{batch}.pptx')
  .option('--locale <code>', 'number and date style, e.g. en-US, en-GB, de-DE', 'en-US')
  .option('--currency <code>', 'currency for the "currency" format, e.g. USD, EUR', 'USD')
  .addOption(new Option('--missing <mode>', 'what to do with tags that have no data').choices(['error', 'keep', 'blank']).default('error'))
  .action((template: string, data: string, opts) => run(() => batch(template, data, opts, consoleIo)));

program
  .command('check')
  .description('Check a template, and optionally a workbook, for problems without writing anything.')
  .argument('<template>', 'tagged PowerPoint template (.pptx)')
  .argument('[data]', 'data workbook (.xlsx) to dry-run against the template')
  .option('--locale <code>', 'number and date style', 'en-US')
  .option('--currency <code>', 'currency for the "currency" format', 'USD')
  .addOption(new Option('--missing <mode>', 'what to do with tags that have no data').choices(['error', 'keep', 'blank']).default('error'))
  .action((template: string, data: string | undefined, opts) => run(() => check(template, data, opts, consoleIo)));

program
  .command('license')
  .description('Show, activate or remove the DeckForge license on this computer (no network needed).')
  .argument('[action]', 'status | activate | remove', 'status')
  .argument('[key]', 'license key, or path to the license file (for activate)')
  .action((action: string, key: string | undefined) => run(() => license(action, key, consoleIo)));

async function run(task: () => Promise<number>): Promise<void> {
  try {
    process.exitCode = await task();
  } catch (e) {
    if (e instanceof UserError) {
      consoleIo.err(`error: ${e.message}`);
      process.exitCode = 1;
      return;
    }
    consoleIo.err('DeckForge hit an unexpected problem. Please report it with the details below.');
    consoleIo.err(e instanceof Error ? (e.stack ?? e.message) : String(e));
    process.exitCode = 70;
  }
}

void program.parseAsync();
