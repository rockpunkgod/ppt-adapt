`ink-replay-timing.xml` is a minimal native Replay timing tree for an original three-trace InkML page. It contains timing behavior and placeholder shape IDs; the builder remaps those IDs to the current page.

The native behavior is preset63 entrance with `drawProgress` 0→1, 5000ms. The older COM interface reports EffectType0; recreating that enum value destroys the Replay behavior. Keep the native behavior and move/rebind it during bookmark polish.
