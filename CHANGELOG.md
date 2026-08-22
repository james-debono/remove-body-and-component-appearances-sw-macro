# Changelog

Semantic Versioning. `MAJOR` reaches 1 when the behaviour is settled enough
to promise not to break it; 0.x is an honest statement that it may still move.

---

Companion to Apply Unique Colours: clears the appearances that macro writes, and nothing else.

## 0.5.4 — 2026-08-21

- Moved to its own repository. The `Source` URL in the header points at the
  new repository.
- No functional change.

## 0.5.3 — 2026-08-20

- The `Source` URL in the header now points at the renamed repository,
  `apply-colours-sw-macro`. No functional change.

## 0.5.2 — 2026-08-13

- The completion dialog reported the previous version number. See the note under
  Apply Unique Colours 0.11.2.

## 0.5.1 — 2026-08-09

- Released under the **MIT licence**, with the full text carried in the code
  itself. Header brought into line with the other macros. No functional change.

## 0.5.0

- **An appearance touching a body is cleared whole**: 0.4.0 dropped a face only when the body it belongs to was in the same appearance, which turns out never to happen. Propagation stamps the *seed* body's colour onto a *different* body's faces, so the appearance holds the seed body alongside faces belonging to the mirrored or patterned copy. Tested with propagation switched back on, the rule never fired once and the mirrored body stayed coloured. In a part the test is now per appearance rather than per entity: one body is enough to claim the whole thing, faces included. An appearance touching no body was applied to a face or feature deliberately and is still left alone, which is what preserves hand-applied colours.
- Assemblies keep the per-entity test. Anything in an assembly's appearance other than a component belongs to a referenced part file and must not be touched.
- **Diagnostics off by default** now the macro is proven. The description no longer claims "face appearances are left alone", which was too broad - most faces in a patterned model carry propagation artefacts rather than anything the user applied.

## 0.4.0

- **Faces now go with their body**: A body's appearance is listed against the body *and* against every face inheriting it. 0.3.0 treated those faces as things to preserve, which was wrong twice over. It left the colour visible - a face-level appearance beats a body-level one, so stripping the body changed nothing on screen - and it put the faces back one `AddEntity` call at a time, 5,359 of them on a 56-body weldment, taking 161 seconds. A face is now dropped when the body it belongs to is being dropped from the same appearance. A face the user coloured deliberately lives in its own appearance and never shares an entity list with a body, so it is still left alone.
- **The rebuild was never the problem**: Instrumenting the run showed `Strip` at 161.74 s against `Rebuild` at 0.75 s. `EditRebuild3` had been the suspect for both the slowness and the stale display; it was responsible for neither.

## 0.3.0

- **Instrumentation**: Per-appearance reporting of what each appearance is attached to, broken down by entity kind, plus counts of re-attach and write-back failures and a timing split across scan, strip and rebuild. Added to locate a fault rather than guess at it; the entity breakdown is what identified the cause immediately.

## 0.2.0

- **Removal scoped to the active display state**: 0.1.0 used `RemoveMaterialProperty`, which is scoped by *configuration*. Several display states live under one configuration, so clearing one cleared them all - the opposite of the behaviour this macro is supposed to mirror. Rewritten around `IModelDocExtension::GetRenderMaterials2` and `AddDisplayStateSpecificRenderMaterial`, which take a display state option instead. An appearance is a render material plus its list of attached entities, and no call detaches a single entity, so the sequence is: read the list, drop all of them, add back the survivors, write the result to this display state.
- **Assembly path fixed**: `IComponent2::RemoveMaterialProperty` exists but not with the argument list 0.1.0 used, which failed to compile. Components are now cleared by detaching them from the appearance, the same mechanism used for bodies, so there is one code path for both document types and no undocumented signature to guess at.
- **Viewport refreshed properly**: `EditRebuild3` replaces `GraphicsRedraw2`. Without it the old colours stayed on screen until the display state was switched away and back, making a successful run look like it had done nothing.

## 0.1.0

- Initial version. Removed body appearances in a part and component appearances in an assembly, at the current configuration.

---
