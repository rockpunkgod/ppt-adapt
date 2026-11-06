# True glyph-hole portal over an automatic video

Backend **hybrid**; design category: spatial reveal / cinematic title. OfficeCLI builds the two-page editable foundation and embeds an original video; native PowerPoint COM subtracts the glyph, adds synchronous exits and 90x Grow/Shrink, and verifies native playback.

Source: masterclass [9Ewb0Rd3P3E](https://www.youtube.com/watch?v=9Ewb0Rd3P3E), 5:09:54–5:17:14, especially 5:13:00 subtraction, 5:15:02 stacking and 5:16:19 smooth-start dialog. Full available transcript and exact source frames were reviewed through Ego. At 5:16:19 the dialog still shows 150%; the subsequent spoken instruction specifies 9000%, so do not miscite that screenshot as showing the final value.

The tested original is a **single H** monogram with Impact (installed locally), purple background, own moving-clock diagnostic MP4. Source Hansome/Singapore footage, radial gradient, first-page Ink Replay and multi-letter 10% staggering are not bundled or claimed as reproduced. The H center stroke is intentional: the growth pivot must sit inside a glyph cutout, otherwise an O counter or letter gap can keep the screen center covered after expansion.

Layer order on page 2: actual MP4 at back → native Subtract compound-path mask → matching white H → caption. Initial foreground H fills the hole. Both foreground objects exit upward over 1 second with maximum smooth start while the mask grows to **9000% / 90x** over 2 seconds, also with maximum smooth start. All share page-load start; page 2 has no transition. The native video starts automatically and remains at its original size. This is not a picture zoom.

Run with fresh output paths:

```
python scripts/glyph_portal_build.py --officecli EXE --video OWN_MP4 --poster OWN_PNG --outdir FRESH_DIR
pwsh -File scripts/glyph_portal_polish.ps1 -InputDeck BASE.pptx -OutputDeck FRESH.pptx -Evidence FRESH_NATIVE_DIR
pwsh -File scripts/glyph_portal_record.ps1 -Deck FRESH.pptx -Evidence FRESH_PLAY_DIR -Ffmpeg EXISTING_FFMPEG
python scripts/glyph_portal_verify.py --deck FRESH.pptx --video OWN_MP4 --report FRESH_JSON
python scripts/glyph_portal_analyze.py --movie FRESH_PLAY_DIR/native.mp4 --ffmpeg EXISTING_FFMPEG --report FRESH_JSON
```

Use the original moving-clock MP4/poster produced by `media_controls_assets.py` for the diagnostic pixel audit. Arbitrary owned video can be used for production but those diagnostic color/pixel criteria do not apply. Native playback is muted (`vol=0`); no audio/device/global preference change. Existing tools are required and not automatically installed. Runtime is Windows PowerPoint with COM, PowerShell7 and OfficeCLI; the recorder uses `media_controls_window.ps1` from the same scripts directory. It refuses another active slideshow and closes only its own document/show.

The recorder starts on static page 1, starts actual desktop capture, then invokes COM Next into page 2; this avoids missing the short effect during slideshow Run startup. It validates automatic animation/layer occlusion, not physical mouse clicking. Source clock and capture clock are distinct; no exact synchronization claim. Acceptance requires true-hole debug PNG, actual intermediate reveal with foreground leaving, eventual full video, and changing clock. XML alone is insufficient.

Research prototype and packaged fresh-forward both passed 11 independent XML/identity checks and 5 real-screen pixel checks. Proof: `verification/glyph-portal-forward/`. Native mask is a 16-node freeform with two closed contours; persisted scale is `to=9000000`, duration2000ms, acceleration100000. Six-second real-screen footage contained 16 partial-reveal samples at 30fps analysis while the H foreground moved upward, then full moving video; native true-hole PNG, initial PNGs and contact sheet were viewed. Initial page 1/page 2 native renders look continuous, but are not pixel-identical: raw RGB RMS0.6703/255, 2424 differing channels, concentrated at glyph-edge rasterization. Do not claim exact pixel identity. No broad claim about all title layouts is made.
