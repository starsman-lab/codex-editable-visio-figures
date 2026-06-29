---
name: visio-local-paper-figure
description: Create, revise, and export editable Microsoft Visio paper figures on a local Windows machine using the installed Visio desktop app. Use when Codex should connect to local Visio, manage a local seed .vsdx template, generate or edit paper figures from a JSON drawing spec, export PNG/SVG/PDF previews, or hand off to Windows-based Visio UI refinement after a scripted draft exists.
---

# Visio Local Paper Figure

## Core Rule

Keep the Visio document as the source of truth. Deliver `.vsdx` first, then export PNG, SVG, or PDF from the saved Visio source.

Prefer a stable local seed `.vsdx` file over creating a document from scratch through COM. On this machine class, template- or scratch-document creation may be less reliable than opening and modifying a known-good `.vsdx` seed.

Do not satisfy a figure-building request by dropping a full reference image onto the page as the final artifact. A reference image is allowed only as a temporary tracing layer that is removed or hidden before delivery.

## Required local asset

This skill works best when a small seed Visio file exists, for example:

- `assets/seed-paper-figure.vsdx`

Use that seed file as the base document for scripted edits. If the seed file is missing, ask the user to create or provide one, or guide them to make a blank one manually in Visio and save it into the skill assets directory.

## Workflow

1. Probe the local environment.
   - Run `scripts/visio_probe.ps1` before the first substantial Visio action in a session.
   - Confirm that Visio COM is available and capture the installed path and version.

2. Choose an execution mode.
   - `spec build`: copy a seed `.vsdx`, then draw or revise a figure using `scripts/visio_apply_spec.ps1`.
   - `export`: export a page from an existing `.vsdx` using `scripts/visio_export.ps1`.
   - `launch + refine`: open Visio and use live UI refinement after a scripted draft exists.

3. For rebuilds, define the figure before drawing.
   - Identify page size, orientation, panel structure, repeated modules, labels, connectors, and final export formats.
   - Build a coarse panel map first, then add internals.
   - When the source is a reference image, convert the image into an object inventory rather than tracing decorative details too early.

4. Use the JSON spec path for deterministic drawing.
   - Read `references/spec-schema.md` when creating or editing a spec.
   - Start with rectangles, ellipses, text blocks, lines, and connectors.
   - Use top-left coordinates in the spec. The script handles the Visio coordinate conversion internally.

5. Use Visio UI refinement only as a second pass.
   - If exact typography, alignment, routing, or ribbon formatting is easier interactively, use Windows-based Visio interaction after the scripted draft exists.
   - Keep UI refinement focused on polish, not on reconstructing the whole figure from scratch when a spec-driven draw would be more reliable.

6. Verify the result.
   - Save the `.vsdx`.
   - Export at least one preview when feasible.
   - Confirm that requested files exist and are non-empty.
   - State clearly whether the output is native editable Visio content or partly manual follow-up work is still needed.

## Scripts

- `scripts/visio_probe.ps1`
  - Detect the Visio executable and test COM automation.
- `scripts/launch_visio.ps1`
  - Open Visio or a target `.vsdx` locally.
- `scripts/visio_apply_spec.ps1`
  - Copy a seed `.vsdx`, then create or update a Visio page from a JSON shape specification.
- `scripts/visio_export.ps1`
  - Export PNG, SVG, and PDF from an existing `.vsdx`.

## References

- Read `references/workflow.md` for end-to-end operating guidance.
- Read `references/spec-schema.md` when generating or modifying a drawing spec.

## Acceptance Criteria

A result is acceptable when:

- the `.vsdx` opens in Visio and remains editable;
- major panels, labels, and flow direction match the intended figure;
- requested PNG, SVG, or PDF outputs were exported from the saved Visio source;
- the final response says what was automated, what seed file was used, and what remains if the figure is not yet publication-ready.
