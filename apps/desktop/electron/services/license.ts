/**
 * License state. L3 ships trial mode only (every deck gets the "Made with DeckForge" footer);
 * L4 replaces this with offline Ed25519 license verification.
 */
import type { LicenseView } from '../api.ts';

export class LicenseService {
  status(): LicenseView {
    return { mode: 'trial', summary: 'Trial: generated decks carry a small "Made with DeckForge" footer.' };
  }
  activate(_key: string): LicenseView {
    return { ...this.status(), problem: 'License activation arrives in the next version.' };
  }
  deactivate(): LicenseView {
    return this.status();
  }
}
