# Workflow

## Recommended production setup

For practical repeated use, keep a known-good seed Visio file under the skill assets directory, for example:

- `assets/seed-paper-figure.vsdx`

That seed file should already open correctly in local Visio. It can be nearly blank; the value is that it is a stable `.vsdx` base for COM edits and exports.

## 1. Decide the job type

### A. Build from a structured idea

1. Confirm the seed `.vsdx` path.
2. Draft a JSON spec.
3. Run `scripts/visio_apply_spec.ps1` with the seed file.
4. Export a PNG preview.
5. Refine the spec or use live Visio UI for small adjustments.

### B. Rebuild from a reference image

1. Decode the image into objects.
2. Create a coarse JSON spec for the large layout first.
3. Build the first-pass figure from the seed `.vsdx`.
4. Use live Visio refinement only for fine alignment, connector routing, or typography.

### C. Export-only

1. Run `scripts/visio_export.ps1` on an existing `.vsdx`.
2. Report the generated files.

## 2. Practical advice

- Do not rely on creating a brand-new Visio file from scratch unless that path has been validated on the current machine.
- Prefer copying a known-good seed `.vsdx` into the workspace and editing the copy.
- If the machine changes, rerun `visio_probe.ps1` and a save/export smoke test before trusting automation.
