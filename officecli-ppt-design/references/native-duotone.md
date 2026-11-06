# Native two-color duotone, with original image bytes preserved

Backend **officecli-native**. Category: picture color / continuous tonal mapping. The recipe attaches native DrawingML `a:duotone` directly to a picture's `a:blip`. It does not bake a recolored bitmap, convert to grayscale plus a colored cover, or require COM for creation.

Source: [9Ewb0Rd3P3E](https://www.youtube.com/watch?v=9Ewb0Rd3P3E),52:51–54:24. Full relevant transcript was read; exact53:00 original photo,53:42 two source endpoint swatchesFF623C/130063 and54:06 ShapeFactory website were inspected through Ego. The author explains selecting a brighter/darker pair and downloading the recolored image from a website. His statement that PowerPoint lacks the feature describes his UI workflow at the time; it is not a restriction on native DrawingML. Native reconstruction here tests the two-endpoint mechanism, not byte-for-byte equality to the website's color pipeline or reuse of its photos.

The [official OpenXML Duotone documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1) specifies per-pixel interpolation between two colors, permits the element under `blip`, and lists Office2007+ availability. It does not justify assuming universal8-bit sRGB/luma arithmetic for every source image. Choose distinct dark/light colors to retain contrast; first child maps the tested black end, second the white end.

For a picture already inserted by OfficeCLI, append the native effect through the supported raw XML command:

```
officecli raw-set OWN.pptx /slide[1] --xpath '//p:pic[p:nvPicPr/p:cNvPr[@name="your-picture"]]/p:blipFill/a:blip' --action append --xml '<a:duotone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="201044"/><a:srgbClr val="F2D787"/></a:duotone>' --json
```

Inspect existing picture effects first; replace an existing duotone deliberately instead of appending duplicates. Do not add a second effect blindly or recolor the stored image as a substitute. Use the CLI's explicit open/close to flush before an external native reader opens the file, then validate the schema.

Diagnostic fixture:

```
python scripts/duotone_build.py --officecli EXE --outdir FRESH_DIR
python scripts/duotone_verify.py --deck FRESH_DIR/duotone.pptx --source FRESH_DIR/original-gray.png --report FRESH_JSON
pwsh -File scripts/duotone_export.ps1 -Deck FRESH_DIR/duotone.pptx -Evidence FRESH_NATIVE_DIR
python scripts/duotone_analyze.py --native FRESH_NATIVE_DIR --ffmpeg EXISTING_EXE --report FRESH_JSON
```

Builder creates an original neutral-gray PNG with16steps, continuous256-level ramp and an original synthetic landscape, then three editable pages: original,201044→F2D787, reversed pair. All image parts must retain the original bytes. The diagnostic fixture's fixed palette/coordinates are not a universal checker for arbitrary photos. Creation needs Python standard library and OfficeCLI only; optional validation uses existing FFmpeg plus Windows native PowerPoint COM, read-only PNG export and no global Office setting change. Output paths must be fresh, no dependency installation.

Research and packaged fresh-forward both passed12independent structure/original-byte checks and10native pixel checks. Proof: `verification/duotone-forward/`; native original/forward/reversed PNGs were inspected visually. Black maps exactly(32,16,68), white to(242,215,135), reversed on the third page; all14interior steps lie between the endpoints, both ramps are monotone and have252sampled colors. Neutral-gray16-step encoded-RGB linear diagnostic error was0in this session; this is a measured result for this source, not a general gamma/color-source formula. Colored-input luminance weights and transparent-pixel compositing were not independently calibrated. Original PNG SHA25699138ebd16f84c2fe525bca6c032a94b2470975c4f74c37503714babcb929dbb stayed unchanged in all three embedded copies. Schema validation passed; native COM was read-only validation only.
