# Media controls

These files contain synthetic format controls. [PROVENANCE.json](PROVENANCE.json) records
fixture hashes, generators, independent oracles and the scope of their evidence. ImageMagick,
LibTIFF, JPEG reference tools, Deark and Apache POI's independent metafile parser are fixture
oracles; they are not production validation dependencies.

The [media owner tests](../../java/dev/erst/gridgrind/engine/media/) exercise framing,
coding and structural relationships. TIFF controls include both byte orders, BigTIFF, strips,
tiles, separate planes, predictors, fax and JPEG layouts. The corpus includes intentionally
malformed controls such as `framed-invalid-lzw.tiff`. Filename labels alone do not
establish validity, and a structural acceptance result is not a rendering certificate.
