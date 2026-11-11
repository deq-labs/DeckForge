# Releasing DeckForge Lite

## Before the first public release (one-time, needs you)

| Item | Why | Where it goes |
| --- | --- | --- |
| Windows code-signing certificate (OV/EV, or Azure Trusted Signing) | Without it SmartScreen warns on install | Repo secrets `WIN_CSC_LINK` (base64 .pfx), `WIN_CSC_KEY_PASSWORD`; or `win.azureSignOptions` in `apps/desktop/electron-builder.yml` |
| Apple Developer account ($99/yr) + Developer ID Application certificate | Gatekeeper blocks unsigned, un-notarized apps | Secrets `MAC_CSC_LINK`, `MAC_CSC_KEY_PASSWORD`, `APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD`, `APPLE_TEAM_ID` |
| Update host (S3/R2 bucket or a folder on your site, over https) | electron-updater downloads from it | Repo variable `DECKFORGE_UPDATE_URL` |
| App icon (1024×1024 PNG) | The default Electron icon is used until then | `apps/desktop/build/icon.png` |
| EULA text | The installer shows `apps/desktop/build/eula.txt` (placeholder now) | same file |
| Production public key + license server URL | See `docs/licensing-operations.md` | `packages/licensing/src/config.ts` |

## Each release

1. Bump `version` in `apps/desktop/package.json` (and `apps/cli/package.json`).
2. Tag and push: `git tag v0.2.0 && git push origin v0.2.0`.
3. The **Release** workflow builds Windows (.exe) and Mac (.dmg + .zip) installers plus the CLI
   bundle, signs and notarizes when the secrets exist, and creates a **draft** GitHub release.
4. Install the drafts on a clean Windows and Mac machine. Generate a deck from each template pack.
5. Upload the installers, `.blockmap` files and `latest.yml` / `latest-mac.yml` to the update
   host, and the installers to your download page. Publish the GitHub release.

The build's release date (the tag date) decides which license keys cover it; see design-log 0007.

## Local builds

```sh
pnpm --filter @deckforge/desktop build
cd apps/desktop
DECKFORGE_UPDATE_URL=https://example.invalid npx electron-builder --win --dir --publish never   # unpacked, fastest
DECKFORGE_UPDATE_URL=https://example.invalid npx electron-builder --win --publish never         # installer
pnpm --filter @deckforge/cli build   # apps/cli/dist/deckforge.cjs, runs with plain `node`
```
