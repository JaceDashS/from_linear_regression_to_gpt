# Workspace Editing Policy

## Existing files are the deliverables

- When the user asks to revise an existing presentation or other artifact, modify the existing file in place.
- Do not create a newly named final deliverable such as `*_updated`, `*_revised`, `*_with_*`, or `*_final` unless the user explicitly requests a separate copy.
- Temporary working copies, extracted assets, renders, and verification files may be created only under `tmp/`.
- Verify the edited result before replacing the existing deliverable. After verification, write the completed result back to the original filename.
- Preserve unrelated user changes and do not replace the original with an incomplete or unverified build.

## Work incrementally

- Do not rebuild, regenerate, re-inspect, or re-render the entire presentation for a small localized change.
- Edit only the affected slides and objects whenever the file format and available tools allow it.
- Verify the affected slides first. Run deck-wide rendering or inspection only when the change can reasonably affect the whole deck or when the user explicitly requests it.
- Reuse existing inspection results, renders, scripts, and extracted assets when they are still current instead of recreating them.
- Keep progress updates concise and avoid loading or printing large inventories unless they are required for the requested change.

## Confirm expensive work first

- If a requested approach is likely to require substantial token usage, long-running full-deck processing, broad research, or repeated generation, tell the user before starting the expensive portion.
- Briefly explain why it may be expensive and propose a smaller incremental approach when possible.
- Do not begin the expensive portion until the user confirms, unless the user already explicitly requested full-deck regeneration or exhaustive review.
