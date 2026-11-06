# Custom monotone-band glyph envelope

Backend **hybrid**: OfficeCLI creates editable text/target foundation; nativePowerPoint converts actual font glyphs to contours; mathematical sampling maps the contours and OfficeCLIwrites nativecustGeom. It is vector authoring, not an external-render recipe. Output points are editable inPowerPoint, with a separate editable original textbackup.

Source: BV1oPHgeiEgb50P P25 792–988.3s. Fullrequested captionsegment and exact918/928/931Ego frames were inspected. Illustrator ellipse→Knife partitions→boldtext→targetBringToFront→Object/EnvelopeDistort/MakewithTopObject(Alt+Ctrl+C) is the actual source operation;931shows thetext in the topcap. This recipe is a restricted reconstruction of contour fitting, **not Adobe's live Envelope object or its internal mapping algorithm**, and not a presetWordArt/textWarp substitute.

Domain: a single vertical band described by strictly increasingx samples spanning0..1, with0<=top(x)<bottom(x)<=1. No crossings, overhangs, targetholes, disconnectedpieces, collapsingtips or arbitrary topology. Source's pointed ellipse cap has zero-height tips and is not claimed as replicated. Source/mapped glyphs are closed contourshapes, not livefontcontent; changingtext/target requires rerun, while page1keeps editabletext.

```
python scripts/envelope_band_build.py --officecli EXE --outdir FRESH_BASE
pwsh -File scripts/envelope_band_outline.ps1 -InputDeck FRESH_BASE/envelope-before.pptx -OutputDeck FRESH_OUTLINED.pptx -Evidence FRESH_LOG
python scripts/envelope_band_warp.py --input FRESH_OUTLINED.pptx --officecli EXE --outdir FRESH_WARP
python scripts/envelope_band_verify.py --deck FRESH_WARP/envelope.pptx --model FRESH_WARP/warp-model.json --report FRESH_JSON
pwsh -File scripts/envelope_band_export.ps1 -Deck FRESH_WARP/envelope.pptx -Evidence FRESH_PNG
python scripts/envelope_band_analyze.py --native FRESH_PNG --ffmpeg EXISTING_EXE --report FRESH_JSON
```

Alloutputs fresh; no dependencies/fontinstallation. Fixture uses locally installedArialBoldBOLD with4holes(B2/O1/D1), two original nonpresetbandboundaries. Nativeoutline usesIntersect of thetextwith a coveringrectangle on its own selected slide; original page1textis preserved. Nativegeometry is read from the resultingfile, never from a researchpptx seed. Sourcecommands Move/Line/Quadratic/Cubic/Close are sampled; unsupported commands/formulas reject rather than silently approximate anotherkind of path.

Optional `--targets OWN.json` onwarp accepts exactlytwo targetprofiles for the fixture's two outputpages:

```
[
  {"name":"own-band-a","samples":[[0,0.2,0.8],[0.5,0.1,0.9],[1,0.3,0.85]]},
  {"name":"own-band-b","samples":[[0,0.1,0.7],[0.5,0.3,0.95],[1,0.2,0.8]]}
]
```

Each row isx/top/bottom. Profiles are validated before creating/opening the outputdeck. Mapping:x=u,y=top(u)+(bottom(u)-top(u))*v; Jacobian=bottom-top>0, a nonfolding continuous map before polygon discretization. Targetandglyph are actualcustomDrawingMLpaths. ThelargeXML is sent via supportedCLIbatchJSONpart/xpath/action/xmlfields to avoidWindowscommandline-length limits; no directZIPrewrite authors thedeck.

For anotherword/font, preserve atextbackup and independently revalidate outlinecommand support, spacing, counterholes, bounds and legibility. The fixedBOLDpixelchecker is a diagnostic, not a universalOCR/qualitychecker. Newprofiles/fonts may need finer sampling or designadjustment; positiveJacobian alone does not establish legibility or discretized topology.

Research and packagedfreshforward passed17structure/modelchecks plus6nativePNGchecks and schema validation. Proof: `verification/envelope-band-forward/`; originalbackup, nativeoutline and bothcustomwarps were actually viewed. Sourceglyph197nodes becomes8closedcontours/4121sampledvertices; both customoutputs preserve8contours. MinimumJacobian0.64/0.501586, everyactualoutputvertex within its suppliedband. No rasterpicture orpresettextWarp. Native sourceoutline andbothwarps each have4connectedlettercomponents and4enclosedcounters, and were visually read asBOLD. Countercounts alone do not prove readability.

Untested: arbitrarytopobject paths, Adobe-liveenvelope editing, zero-heighttips, concaveoverhangs, differenttext/font/targets, exactAdobealgorithm/appearance equivalence. Do not mark the wholearbitrary-envelope familycomplete from this bounded subrecipe.
