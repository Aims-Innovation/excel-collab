# Excel-Collab Fork Integration — Plan & State

_Living document tracking the port of `nusr/excel` (excel-collab) into
`Aims-Innovation/excel-collab` and its consumption from `web-consumer`._

## TL;DR — where we are

| Layer | Status |
|---|---|
| `web-consumer` on branch `AzharGhani/consume-excel-collab-fork` | Points at fork tag `release/document-editor-v0.1.13.2`. React + react-dom pinned to 19.1.0 (resolutions + overrides). `optimizePackageImports` no longer routes React through Next's vendored canary. PDFViewer lazy-loaded + `transpilePackages`. Stale `patches/` dir + `patch-package` hook removed. |
| `Aims-Innovation/excel-collab` on branch `document-editor` | Clean flatten (monorepo stripped to `excel-collab` sub-package). Vite externalizes `react` / `react-dom` + subpaths so react-dom stays in the consumer. **No yally customizations ported yet** — this is still the vanilla upstream shape, rebuilt. |

## What works after the switch

- DOCX opens and edits fine. No regression.
- XLSX editor mounts. Chrome renders. Typing into cells works and persists via REST auto-save.
- The React "Incompatible versions" throw is gone.
- DataCloneError from the old Worker-render path no longer fires (fork builds clean; the error appeared to be version-specific to the npm 0.1.13 bundle).
- WebSocket falls back to LongPolling (infra — not our concern).

## What's broken or missing

### Blockers

1. **XLSX original file content not rendered in the grid.**
   - Editor opens empty even for freshly uploaded xlsx files (so it's NOT stale backend state).
   - Typing works → controller is alive + Y.Doc is bound.
   - The seed-from-blob path (`useInitialWorkbookSeed` in
     `src/app/documents/document-editor/components/spreadsheet/ExcelCollabSystem.tsx`)
     never puts the file's rows into the grid.
   - Possible causes (need logging to confirm):
     - Download blob never arrives → `fileBlob` stays null → effect returns early waiting for it.
     - Backend's batch reply is non-empty → `isFirstUser=false` → seed is skipped.
     - SheetJS `sheet_to_json` returns empty rows for this particular file shape.
     - `controller.setCell` throws silently.
   - **Next action:** upgrade the `logger.debug` calls in `useInitialWorkbookSeed` + the download effect to `logger.info` (or add a banner) so we can see what's happening. See "Debugging steps" below.

2. **PDF still throws `Object.defineProperty called on non-object` in dev.**
   - `transpilePackages: ['pdfjs-dist', 'react-pdf']` pushed in `7c8f6459`
     — but `next.config.js` is read once at dev-server start; full
     `yarn dev` restart required (not just `.next` purge).
   - If the error persists after a real server restart, we may need
     extra webpack config (known Next 15 regression around mixed ESM +
     `output: 'export'`).

### Cosmetic / low-prio

3. **`Yjs was already imported` warning** — excel-collab bundles its own
   vendored yjs, project has its own. Known, non-blocking. Optional
   fork patch: strip yjs's "already imported" check in the fork's
   bundled copy.

4. **`No client method with the name 'currentusers' / 'editorpermission'`
   warnings** — the backend emits these SignalR events; we don't
   subscribe. Cosmetic. There's an unmerged PR `#1581` (branch
   `AzharGhani/fix-excel-collab-theme-and-strictmode`) that silences
   them. Needs to be rebased onto the fork-consumption branch or
   merged separately.

5. **Menubar chrome (New File / Rename entries, user badge, theme
   toggle)** — not yet hidden by fork; still masked by CSS in
   `ExcelCollabSystem.css` (unchanged). Waiting on the fork patch.

6. **Locale sync** — `window.__yallyExcelI18n` not exposed by the
   fork baseline. Our `ExcelCollabSystem.tsx` call is guarded by
   optional chaining (`?.changeLanguage`), so it silently no-ops.
   Toolbar/menu labels don't follow app locale until fork patch is ported.

## Patch-porting roadmap (to land on fork's `document-editor` branch)

Each becomes one TypeScript commit on the fork. Ordered by impact.
Once all are in, rebuild `lib/`, tag `release/document-editor-v0.1.13.3`
(or similar), bump the web-consumer pin.

| # | Change | Fork source area (approximate) | Priority |
|---|---|---|---|
| 1 | ~~**Guard Worker.postMessage against DataCloneError.** Wrap the canvas worker's postMessage in the fork source so a function-bearing payload doesn't kill the render. Originally the "option A" runtime guard in web-consumer; cleaner at source.~~ **Landed** on `claude/restart-dev-server-ufgAZ`: new `src/controller/workerGuard.ts` monkey-patches `worker.postMessage` at `initController`. Fast path = direct passthrough; `DataCloneError` → sanitize (drop functions/symbols/circulars, preserve Map/Set/Blob/typed arrays) + retry once with the original transfer list. One-shot `console.warn` on first occurrence. Unit test at `src/controller/__tests__/workerGuard.test.ts`. | `src/controller/` (worker boundary) | High — catches a class of bugs going forward |
| 2 | ~~**Expose i18n singleton on `window.__yallyExcelI18n`.** One-line add in `src/i18n/index.ts` wrapping the existing `Ce = Uz()` export: `if (typeof window !== 'undefined') { (window as any).__yallyExcelI18n = Ce; }`. Unblocks locale sync from the host app.~~ **Landed**: `src/i18n/index.ts` now exposes the `i18n` singleton on `window.__yallyExcelI18n` (with SSR-safe `typeof window` guard). Host app can call `window.__yallyExcelI18n?.changeLanguage('en-US')` to sync locale; the singleton's `.t`, `.current`, `.init` are also reachable for diagnostics. | `src/i18n/index.ts` | Medium |
| 3 | ~~**Drop `location.reload` on language change.** In the library's language picker component (`src/components/…/LanguageSwitcher.tsx` or similar). Remove the `location.reload()` call; the i18n store re-renders reactively.~~ **Landed**: i18n store was NOT reactive — lots of consumers call `i18n.t(...)` during render of `memo`'d components, so a plain reload was the sledgehammer. Fix adds `subscribe(listener)` to the singleton, wires `useSyncExternalStore` into `Excel` root (`src/editor.tsx`), and passes `key={language}` to `ExcelEditor` — language change now triggers a clean remount of the editor subtree (preserves Y.Doc / controller / backend state; only React UI resets). `location.reload()` removed from `src/containers/MenuBar/I18N.tsx`. RTL attribute update moved to a language-keyed effect so it follows switches too. | `src/i18n/index.ts`, `src/editor.tsx`, `src/containers/MenuBar/I18N.tsx` | Medium |
| 4 | ~~**Remove "New File" / "Rename" File-menu entries.** Host-app owns document identity; these orphan entries confuse users. Source-level removal is cleaner than our current CSS masking. Ideally make it a prop (`hideNewFile?: boolean`) so upstream could accept it.~~ **Landed** as props `hideNewFile` + `hideRenameFile` on `EditorProps` (and therefore `<Excel>`), both defaulting to `false` (upstream behavior preserved). When `hideRenameFile` is set the clickable filename widget on the left also hides, since it's the other entry point into the rename dialog. Web-consumer can now drop its CSS mask by passing `<Excel hideNewFile hideRenameFile … />`. Also fixed a stale `testId="menubar-new-excel"` collision on the rename item → `menubar-rename-excel`. | `src/containers/index.tsx`, `src/containers/MenuBar/index.tsx` | Medium |
| 5 | **StrictMode-safe TB bootstrap.** `TB()` function (container bootstrap) has an empty-deps `useEffect` that races under React StrictMode — add a per-mount useRef guard. | `src/containers/…/TB.tsx` (approximate) | Low-medium (latent; may or may not repro on current consumer config) |
| 6 | **React version check bypass — most likely NOT needed.** Our fork now externalizes react + react-dom subpaths, so the check runs in web-consumer's own react-dom (now aligned at 19.1.0). If the pin ever drifts again we'll revisit. | N/A | Skip unless repros |

## Web-consumer action items

- **Merge / rebase PR #1581** (`AzharGhani/fix-excel-collab-theme-and-strictmode`) into the fork-consumption branch. It adds CurrentUsers / EditorPermission no-op subscriptions, distinct `offline` phase, and tear-down for the SignalR service on failure. Nothing in it fights the fork; it stacks cleanly.
- **Add logging** in `useInitialWorkbookSeed` to debug the blank-grid issue. See "Debugging steps".
- **Verify PDF fix** after full dev-server restart.
- **Keep an eye on** the `Yjs was already imported` warning — if editing ever produces weird replication bugs, we'll need to force single-yjs-copy (hard without full vendor).

## Debugging steps for blank-grid (next session)

Patch `ExcelCollabSystem.tsx` to surface seed-path state:

```ts
// In the blob-download effect (around line ~165):
downloadDocument(numericId).unwrap()
  .then(blob => {
    if (cancelled) return;
    logger.info('[ExcelCollabSystem] blob downloaded', { size: blob.size, type: blob.type });
    setFileBlob(blob);
  });

// In useInitialWorkbookSeed:
React.useEffect(() => {
  logger.info('[useInitialWorkbookSeed] effect run', {
    hasController: !!controller,
    hasFileBlob: !!fileBlob,
    isFirstUser,
    seeded: seededRef.current,
  });
  // … existing body
}, [controller, isFirstUser, fileBlob, fileName, kind]);

// Inside seedControllerFromBlob, log sheet counts / row counts:
logger.info('[seedControllerFromBlob] parsed', {
  sheetNames: workbook.SheetNames,
  firstSheetRowCount: /* rows.length */,
});
```

Open a fresh xlsx, check the console — we'll know in one render whether
blob download failed, isFirstUser was false, or SheetJS returned empty.

## Known repo-level breakage (not caused by any patch)

- **`yarn test` does not run.** The flatten commit (`d77a276 feat: Buil`)
  dropped `scripts/jest.config.js` + `scripts/jest.setup.js` +
  `scripts/css-transform.js` + `scripts/css-mock.js` without relocating
  them to the flat layout. Every existing test under `src/**/__tests__/`
  fails to transform (Babel parser chokes on `: Type` syntax because no
  `ts-jest` preset is wired up). This is pre-existing; patch #1's new
  test file (`src/controller/__tests__/workerGuard.test.ts`) is written
  against the original ts-jest + jsdom setup and will run once the config
  is restored. Suggested follow-up commit: restore `jest.config.js` +
  `jest.setup.js` at repo root with paths updated for the flat layout.
  Out of scope for the five patches in the roadmap above.

## Session handover

- **This session (web-consumer)** — finish plan.md, push, continue
  handling web-consumer-side tweaks + PRs.
- **New session (excel-collab fork)** — start with `cd C:\excel-collab`
  and "spawn Claude Code here". Use this plan.md as the working context,
  port patches 1–5 above as individual TypeScript commits, rebuild
  `lib/`, tag next release.
- **Cross-session coordination** — update this plan.md at the end of
  each chunk of work with what landed + what's remaining.

## Links

- Fork: https://github.com/Aims-Innovation/excel-collab
- Fork branch: `document-editor`
- Fork latest tag: `release/document-editor-v0.1.13.2`
- web-consumer branch: `AzharGhani/consume-excel-collab-fork`
- Unmerged helper PR: #1581 (theme + StrictMode + `PermissionDenied` listener)
- Backend hub PR (verified contract): `backend-centralized` PR #785
