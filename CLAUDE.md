# SML Asset Transfer System — notes for future work

Google Apps Script web app: `Code.gs` (backend, `doGet`/`doPost` action dispatch) +
`index.html` (single-page vanilla-JS frontend). Data lives in a bound Google Sheet.

## Deployment is manual — git merge does NOT deploy

Merging a PR only updates the GitHub repo. The live web app only picks up changes
after someone manually copies `Code.gs`/`index.html` into the Apps Script editor and
creates a new deployment version (Deploy > Manage deployments > Edit > New version >
Deploy). Always remind the user of this after merging.

## Sheet schema changes require re-running `setup()` — and even then, order isn't guaranteed

`ensureSheet_()` only runs inside `setup()`, which the user must trigger manually from
the Apps Script editor. It is **not** called automatically per-request. So after adding
a field to any `HEADERS.*` array, an already-deployed sheet keeps its old columns until
`setup()` is re-run — and even then, `ensureSheet_` appends any *missing* header to the
**end** of the sheet's existing columns, not at the position it holds in the `HEADERS.*`
constant.

This has caused real bugs twice in this project (Users sheet `Departments` column,
TransferQueue sheet `Purpose` column): code that wrote new rows as a positional array
literal (e.g. `sh.appendRow([a, b, c])`) assumed the physical column order always
matched the `HEADERS.*` constant order. It doesn't, once a field is inserted anywhere
but the very end of an already-shipped `HEADERS.*` array.

**Rule going forward: never write a new row as a positional array literal that assumes
`HEADERS.*` order.** Always build rows from the sheet's actual live header row:

```js
const values = sh.getDataRange().getValues();
const headers = values[0];
const idx = indexMap_(headers);
const newRow = headers.map(() => '');
newRow[idx.SomeField] = someValue; // guard with `if (idx.SomeField !== undefined)` if the
                                    // column might not exist yet (setup() not re-run)
sh.appendRow(newRow);
```

Existing helpers already follow this pattern: `adminSaveAsset_`, `adminSaveUser_`,
`adminSyncFromSource_`, `addToAssetQueue_`. Keep any new row-insert code consistent
with it. When a write depends on a column that might not exist yet, guard with
`idx.Field !== undefined` and fail with a clear "please re-run setup()" error rather
than silently misfiling data into the wrong column.

## AssetStatus is computed live; Tag is persisted

`AssetStatus` (`Active`/`Sold`/`WrittenOff`) is never stored — `getDisposedAssetStatus_()`
recomputes it on every read by scanning approved Sale/WriteOff documents. This is
intentional (always accurate, no sync step). But it means the raw `Assets` sheet itself
never visibly shows "sold" — only the app's computed views do.

To keep the raw sheet and the Dashboard's tag-based breakdown in sync with reality,
`setAssetsTag_(assetIds, tagValue)` writes `Tag='ขาย'`/`'ชำรุด'`/`'ใช้งาน'` onto the
Assets sheet whenever a Sale/WriteOff gets approved or an asset gets restored
(`createSale_`/`createWriteOff_` auto-approve paths, `decideSale_`/`decideWriteOff_`,
`adminRestoreAsset_`). If a new disposal/restore path is added, wire it in too.

## Verification checklist before every push (no live test environment available)

1. `node --check` on `Code.gs` and on the extracted `<script>` block of `index.html`.
2. Every frontend `getElementById('...')` has a matching `id="..."` literal somewhere
   in the file.
3. Every frontend `action:'...'` string has a matching backend `case '...':` in the
   `doGet`/`doPost` switch.
4. For visual/print changes, render with headless Playwright
   (`/opt/pw-browsers/chromium`) and screenshot rather than guessing.

## Standing workflow rules

- Work on branch `claude/admin-settings-appsheet-viewdata-k1n1pz`. If it was already
  merged into `main`, restart it from `main` before adding new commits (don't stack on
  merged history).
- Only create/merge a PR when the user explicitly says "merge" — never preemptively.
- Queues (Transfer/Sale/WriteOff) share one sheet (`TransferQueue`) distinguished by a
  `Purpose` column; blank `Purpose` is treated as `Transfer` for backward compatibility.
- Disposing an asset (Sale/WriteOff approval) purges it from *all three* pending queues
  via `purgeAssetFromAllQueues_`, not just the queue the document was created from.
