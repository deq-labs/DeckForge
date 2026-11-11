import type { DeckforgeApi } from '../electron/api.ts';

declare global {
  interface Window {
    deckforge: DeckforgeApi;
  }
}
