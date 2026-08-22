# Remove Body and Component Appearances — development notes

The API findings behind this macro. `README.md` covers what it does and how to
use it.

## What it clears

Body-level appearances in a part, and component-level appearances in an
assembly — the level [Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro) writes to.
Face and feature appearances are deliberately left alone, which is what makes
this the targeted undo rather than a blunt clear.

Everything is scoped to the **active display state**, through `DisplayStateSetting`.
In an assembly the referenced part files are never opened or modified.

## Propagate visual properties — read this before debugging appearances

Pattern and mirror features have a **"Propagate visual properties"** option, on by
default. When on, regenerating the feature stamps the seed body's colour onto the
derived bodies' **faces**. A face appearance beats a body appearance, so colour can
reappear after a rebuild even though removal succeeded. Check that checkbox before
suspecting the macro.

## Dead ends

Each was implemented, measured and rejected. **Do not revisit without new
evidence.**

- **`RemoveMaterialProperty`** (0.1.0). Scoped by *configuration*, and several
  display states live under one configuration, so it cleared them all.
- **Keeping all faces during removal** (0.3.0). Left bodies still coloured, since
  a face appearance beats a body one, and took 161 seconds re-attaching 5,359
  faces.
- **Dropping a face only when its own body is in the same appearance** (0.4.0).
  Never fires. Propagation stamps the *seed* body's colour onto a *different*
  body's faces, so the appearance holds the seed body alongside faces belonging to
  the derived copy.

## Measured performance

Removal runs in about **1 second** on the 56-body reference weldment.

## Known limitations

- **No undo.** Appearance changes leave no undo record.
- Only the active display state is affected.
- In an assembly, only component appearances are cleared; appearances living
  inside the referenced parts are untouched.

## Verification status

Confirmed working in SOLIDWORKS on the reference weldment.

## There is no build step

A `.swp` is a binary VBA project. Editing the `.vba` in `src\` changes nothing
that runs until the source is pasted into the SOLIDWORKS VBA editor and saved.
Treat the `.vba` as the source of truth and re-paste after every change.

Do not patch a `.swp` directly: it stores compiled p-code ahead of the source
text and VBA runs the p-code, so a patched file shows new code in the editor
while still running the old.
