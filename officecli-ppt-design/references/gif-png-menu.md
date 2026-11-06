# Active GIF / inactive PNG menu pairing

Backend: **officecli-native** for GIF/PNG pairing, links, Morph and native GIF presentation. Original diagnostic asset generation (standard-library PNG drawing + existing FFmpeg GIF encoding) is material preparation, not a separate external-render recipe or a fourth-class capability claim. PowerPoint COM/physical input is used only to test the native player. Category: interactive menu state and media lifecycle. This is genuine multi-frame GIF picture media, not video and not a flattened preview.

Source [9Ewb0Rd3P3E](https://www.youtube.com/watch?v=9Ewb0Rd3P3E),7:00:00–7:14:54. Full available transcript was read, six distributed frames and exact7:10:01/7:11:36/7:12:36/7:13:43 operation frames inspected through Ego. Source uses GIF for the active icon, same-source PNG at75% transparency for inactive icons, same `!!iconN` name across picture variants, .7s Morph and one page per selected state. Linked circles are brought in front of icons so pictures do not intercept pointer input. The source export dialog has transparent GIF,400px,0ms extra delay,infinite loop. It does not promise a quantified native restart phase.

Original assets: procedural house/globe/pencil white line drawings with a cyan phase dot,12frames×200ms=2.4s, transparent infinite GIF. Inactive PNG derives from the same phase0 withalpha25%. No sponsor icon files, stock photos, author presentation, or external graphics dependency are bundled. The fourth/initial START page is diagnostic instrumentation so first GIF entry occurs after recording begins.

```
python scripts/gif_menu_assets.py --ffmpeg EXISTING_EXE --outdir FRESH_ASSETS
python scripts/gif_menu_build.py --officecli EXE --assets FRESH_ASSETS --outdir FRESH_DECK_DIR
python scripts/gif_menu_verify.py --deck FRESH_DECK_DIR/gif-menu.pptx --assets FRESH_ASSETS --report FRESH_JSON
pwsh -File scripts/gif_menu_test.ps1 -Deck FRESH_DECK_DIR/gif-menu.pptx -Evidence FRESH_PLAY_DIR -Ffmpeg EXISTING_EXE
python scripts/gif_menu_analyze.py --evidence FRESH_PLAY_DIR --ffmpeg EXISTING_EXE --report FRESH_JSON
```

All outputs must be fresh; no dependency is automatically installed. Creation uses Python standard library, existing FFmpeg and OfficeCLI. Native testing needs Windows PowerPoint COM/PowerShell7, exclusive slideshow/foreground control, and physical mouse movement/clicks. It derives physical client/DWM coordinates, checks actual cursor position and owner window before clicking, and closes only its own document/show. No Office preference, default audio setting, or microphone capture is changed. `-ProbeOnly` records pointer diagnostics and exits intentionally before clicking; it is not a passed test.

Create all target pages before adding links: OfficeCLI rejects forward links to nonexistent slides. Each menu page contains exactly one active GIF and two same-source PNGs. Explicit stable names pair each variant through Morph; selected circle/label/arrow reflects the page state. Linked circle shapes sit above the pictures. Do not infer dynamic state from label text alone.

Prototype independent audit passed18checks: exact original embedded bytes, three real12-frame transparent infinite GIFs, picture/hit stacking and links. Actual20s desktop video with six physical clicks A/B/C/A/B/C passed20state/motion checks: selected GIF animates, two inactive PNGs stay still, only selected icon has a bright moving phase dot, links enter the expected page; the last GIF visibly wraps11→0. Restart phase was audited separately: relative phase origins on revisit differed0.196/0.136/0.147s from first entry (200ms GIF frames), consistent with reset within about one frame in this session. This is **not** an exact frame0/start-time guarantee or a cross-player promise.

An initial pointer guard failed before any click and remains negative evidence; a subsequent no-click diagnostic had expected=actual physical(3837,1684), so the earlier cause is unknown. The passing full run kept the original strict guard; it did not bypass it with COM navigation or relaxed tolerances. Packaged fresh-forward also passed18structural checks,20state/motion checks and schema validation; actual contact sheet/recording was reviewed. Fresh reentry phase differences0.140/0.153/0.159s are likewise within the200ms frame interval, with the same limited interpretation. Proof: `verification/gif-menu-forward/`.

Untested: self-linking the current selected circle, entire-show stop/restart, other GIF disposal/delay encodings, external GIFs, web/mobile players. Reentering a selected page, activity/static pairing and exact phase restart are separate conditions; report them separately.
