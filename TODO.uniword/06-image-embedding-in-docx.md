# Phase 6: Image Embedding in DOCX

## Goal
Convert html2doc's image handling (resize, rename to UUIDs, copy to `_files/` directory) to Uniword's DOCX image embedding model (binary data in `word/media/`, relationships in `word/_rels/document.xml.rels`, content types in `[Content_Types].xml`).

## Background
html2doc's current image pipeline (`mime.rb`):
1. `image_cleanup` — finds `<img>` and `<v:imagedata>` elements
2. `rename_image` — copies local images to `{filename}_files/` with UUID names
3. `image_resize` — uses Vectory to fit images within page dimensions
4. `mime_package` — replaces `src` with `cid:` references, base64-encodes into MHT

Uniword's image pipeline:
1. `ImageBuilder.register_image(document, path)` — reads binary, stores in `DocumentRoot.image_parts`
2. `ImageBuilder.create_drawing(...)` — builds `Drawing > Inline > Graphic > ... > Picture` chain
3. `DocxPackage.to_zip_content` — writes binary to `word/media/`, adds rels and content types

The key difference: html2doc uses Vectory for sizing; Uniword reads dimensions from PNG/JPEG headers and converts to EMU (1 inch = 914400 EMU).

## Tasks

- [ ] Implement `DocxConverter#convert_images(docxml, document)`:
  - Walk all `<img>` and `<v:imagedata>` elements in the cleaned document
  - For each local image:
    1. Get dimensions: use existing Vectory sizing OR Uniword's `ImageBuilder.read_dimensions`
    2. Register with Uniword: `ImageBuilder.register_image(document, image_path)`
    3. Create drawing: `ImageBuilder.create_drawing(document, path, width:, height:, alt_text:)`
    4. Replace the `<img>` element with a placeholder that will become a `Run` containing the `Drawing`
  - For external images (`http://`, `https://`): leave as-is (not embedded in DOCX)
  - For data URIs (`data:image/...;base64`): decode, save to temp file, register as local image

- [ ] Handle image sizing:
  - html2doc resizes to fit page dimensions using `page_dimensions()` and Vectory
  - Convert html2doc's px-based dimensions to EMU for Uniword:
    ```
    1 px = 9525 EMU (at 96 dpi)
    1 pt = 12700 EMU
    1 inch = 914400 EMU
    ```
  - Respect landscape vs portrait orientation (check `@landscape` classes)
  - Handle images within table cells (divide available width by image count)

- [ ] Handle Vectory integration:
  - Option A: Keep using Vectory for dimension calculation, pass results to Uniword
  - Option B: Use Uniword's native dimension handling
  - **Recommend: Option A** — Vectory already handles the edge cases (page margins, landscape)

- [ ] Support image formats:
  - PNG, JPEG, GIF — standard, directly embeddable
  - BMP, TIFF — supported by DOCX
  - EMF/WMF — common in Word documents, embeddable
  - SVG — NOT supported in DOCX (same limitation as MHT)
  - Use `Marcel` for MIME type detection (already a dependency)

- [ ] Handle alt text:
  - Extract `alt` attribute from `<img>` → `docPr(name:, descr:)` in OOXML
  - Default alt text: filename if not provided

## Key Files to Create/Modify
- `lib/html2doc/docx_converter.rb` — image conversion methods
- `lib/html2doc/base.rb` — image_cleanup still runs for DOCX path (provides dimensions)

## Success Criteria
- Local images embedded in DOCX `word/media/` with correct content types
- Image dimensions preserve aspect ratio within page boundaries
- Alt text preserved in DOCX
- Table cell images sized relative to cell width
- External/data URI images handled gracefully (skipped or decoded)
