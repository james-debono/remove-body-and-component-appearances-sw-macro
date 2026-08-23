# Remove Body and Component Appearances

A SOLIDWORKS macro that clears appearances applied at body level in a part, or at
component level in an assembly, and leaves everything else alone.

Works with SOLIDWORKS 2022, 2024 and 2025.

## What it does

This is the targeted removal. It clears the appearances that sit on **bodies** and
**components** — the level that
[Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro)
writes to — so it is the practical undo for that macro.

What it deliberately leaves alone:

- Colours applied to an individual **face**
- Colours applied to a **feature** on its own
- Every display state other than the active one
- Referenced part files, when run in an assembly

That last point matters: an assembly-level run never modifies the parts it
references, so nothing is changed on disk outside the document you have open.

If you want everything gone rather than just the body and component level, use
[Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro)
instead.

## Install

**The macro on its own:** download `Remove-Body-and-Component-Appearances.swp` from
the [latest release](../../releases/latest), then run it with
**Tools > Macro > Run**, or add it to a toolbar with
**Tools > Customize > Commands > Macro**.

**With [MacroShelf](https://github.com/james-debono/macroshelf-sw-addin):** get the
[MacroShelf Collection](https://github.com/james-debono/macroshelf-collection-sw-macro-library/releases/latest),
which packages this macro with its icon and hover text alongside every other macro
in the set. Point MacroShelf at the unzipped folder and it appears as a button.

## Using it

Open a part or assembly and run the macro. It reports how many appearances it
cleared when it finishes.

There is **no confirmation prompt**, because the scope is narrow enough that the
result is predictable — it removes the level these macros write to and nothing
else. The broader Remove All Appearances does ask, for the opposite reason.

## Worth knowing

**Pattern and mirror features with "Propagate visual properties" switched on will
put colour back.** Regenerating such a feature stamps the seed body's colour onto
the derived bodies' faces, and a face appearance is not what this macro clears. If
colour reappears after a rebuild, that checkbox is why.

## Limitations

- **No undo.** Appearance changes leave no undo record.
- Only the active display state is affected. Other display states keep their
  appearances.
- Face and feature appearances survive by design. That is the point of this macro
  rather than a shortcoming, but it does mean a stubborn colour may be at a level
  this macro does not touch.

## Related macros

- [Apply Unique Colours](https://github.com/james-debono/apply-unique-colours-sw-macro)
  — colours every geometrically unique body
- [Remove All Appearances](https://github.com/james-debono/remove-all-appearances-sw-macro)
  — clears every appearance, including face and feature level

## Building from source

`src\Remove-Body-and-Component-Appearances.vba` is the readable source. A `.swp` is
a binary VBA project, so it can only be produced from inside SOLIDWORKS — there is
no build step. Open the `.swp` via **Tools > Macro > Edit**, paste the source in,
and save.

Technical detail is in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licence

MIT — see [LICENSE](LICENSE). Free to use, modify and share. The full licence text
is also carried inside the macro itself, so a `.swp` passed on by itself still
carries its licence.

Created by James Debono, with AI assistance. Everything here was tested by
hand in SOLIDWORKS — nothing that touches the API can be verified any other way.

## Trademarks

SOLIDWORKS is a registered trademark of Dassault Systèmes SolidWorks Corporation.
This project is independent: it is not affiliated with, endorsed by, or sponsored
by Dassault Systèmes, and uses only the published SOLIDWORKS API.
