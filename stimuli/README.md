The experiment discovers the Brady et al. (2008) object images recursively from `stimuli/OBJECTSALL/` at runtime. The current dataset is a flat folder containing 2,400 `.jpg` files; non-image files are ignored.

The experiment automatically reserves separate images for four practice trials, 64 unique Phase 1 targets, 192 unique Phase 1 distractors, and 32 untouched Phase 2 NEW foils. No individual trials need to be defined manually.

## Optional Remote Data Saving

Local Excel export remains enabled. Remote saving is disabled by default in `experiment.js` in the `REMOTE_DATA_CONFIG` section. The current `saveDataRemotely()` function is an integration point only; setting `enabled` does not upload data yet.

To implement DataPipe or another centralized destination, the researcher will need to provide the DataPipe experiment/project ID (or the destination's equivalent), choose the upload mechanism and expected data format, and confirm the service's privacy/consent requirements. If the service requires an API secret or credential, keep it on a trusted server and provide a server-side upload endpoint; do not put secrets in this public browser code or commit them to GitHub. Then implement the request inside `saveDataRemotely()` and test a complete upload before enabling it.
