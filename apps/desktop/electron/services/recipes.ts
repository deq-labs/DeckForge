/** Saved jobs ("recipes") for one-click monthly reruns, kept in a JSON file in the user's app data. */
import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { Recipe } from '../api.ts';

interface RecipeFile {
  version: 1;
  recipes: Recipe[];
}

export class RecipeStore {
  private readonly file: string;

  constructor(dir: string) {
    this.file = join(dir, 'recipes.json');
  }

  async list(): Promise<Recipe[]> {
    try {
      const parsed = JSON.parse(await readFile(this.file, 'utf8')) as Partial<RecipeFile>;
      return Array.isArray(parsed.recipes) ? parsed.recipes.filter(isRecipe) : [];
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') return [];
      // A damaged file shouldn't stop the app; keep it aside and start fresh.
      await rename(this.file, `${this.file}.damaged-${Date.now()}`).catch(() => undefined);
      return [];
    }
  }

  /** Inserts or replaces by id; most recently saved first. */
  async save(recipe: Recipe): Promise<Recipe[]> {
    if (!isRecipe(recipe)) throw new Error('That recipe is incomplete.');
    const rest = (await this.list()).filter((r) => r.id !== recipe.id);
    const next = [recipe, ...rest];
    await this.write(next);
    return next;
  }

  async remove(id: string): Promise<Recipe[]> {
    const next = (await this.list()).filter((r) => r.id !== id);
    await this.write(next);
    return next;
  }

  private async write(recipes: Recipe[]): Promise<void> {
    await mkdir(join(this.file, '..'), { recursive: true });
    const tmp = `${this.file}.tmp`;
    const body: RecipeFile = { version: 1, recipes };
    await writeFile(tmp, JSON.stringify(body, null, 2));
    await rename(tmp, this.file);
  }
}

function isRecipe(r: unknown): r is Recipe {
  if (!r || typeof r !== 'object') return false;
  const x = r as Recipe;
  return (
    typeof x.id === 'string' &&
    typeof x.name === 'string' &&
    !!x.spec &&
    typeof x.spec.templatePath === 'string' &&
    typeof x.spec.dataPath === 'string' &&
    typeof x.spec.outDir === 'string' &&
    !!x.spec.options
  );
}
