# 3. Images not embedded in DOCX output

## Problem

The HTML `rice.html` contains 5 `<img>` elements. The generated DOCX has **no `word/media/` directory** — zero images embedded. The `<img>` elements likely produce empty paragraphs or are silently dropped.

## Root cause (multiple factors)

### 1. `skip_images: true` in cleanup

`process_docx` (base.rb:32) calls `cleanup(docxml, skip_footnotes: true, skip_images: true)`, which skips the `image_cleanup` step. In the legacy MHT path, `image_cleanup`:
- Resizes images to fit page dimensions
- Renames files to UUIDs
- Sets `width`/`height` attributes on `<img>` elements

Without this step, `<img>` elements keep their original `src` (e.g. `rice_images/rice_image1.png`) and have **no `width`/`height` attributes**.

### 2. Image path resolution may fail silently

`DocxConverter#convert_image` (line 669-710) tries to resolve the image path via `resolve_image_path`:
```ruby
image_path = resolve_image_path(src)
unless image_path && File.exist?(image_path)
  para.runs << create_run("[image: #{src}]")
  return
end
```

It tries `File.join(@imagedir, src)` and `File.join(File.dirname(@filename), src)`. If the relative path `rice_images/rice_image1.png` doesn't resolve correctly, it falls back to a `[image: ...]` text placeholder.

### 3. No width/height for ImageBuilder

Even if the path resolves, `width_px` and `height_px` are nil (no resizing happened), so `width_emu` and `height_emu` are nil. The `Uniword::Builder::ImageBuilder.create_run` needs to handle nil dimensions (should read actual image dimensions).

### 4. Images may be inside unconverted elements

Some images may be inside `<div>` children that get flattened. The `convert_image` method exists and is called from `convert_element` when the child is an `<img>`, but only if the element reaches that code path.

## Required changes

1. **Add DOCX-specific image handling** to `process_docx` — either:
   - Run a lightweight image processing step (resize, add width/height attrs) without the MHT-specific UUID renaming and file copying
   - Or have `convert_image` read actual image dimensions when HTML attrs are missing

2. **Fix path resolution** — verify that `File.join("spec/examples", "rice_images/rice_image1.png")` resolves correctly and that the `ImageBuilder` can read the file.

3. **Handle nil dimensions** — `ImageBuilder.create_run` should read actual image dimensions from the file when width/height are not provided.

4. **Image relationship registration** — verify that `ImageBuilder.create_run` correctly registers the image in the Uniword document model so it gets packaged into `word/media/` and the relationship is added to `document.xml.rels`.

## Scope in rice.html

```
1362:  <img src="rice_images/rice_image1.png" />  (Annex A - sample divider diagram)
1481:  <img src="rice_images/rice_image2.png" />  (Annex C - gelatinization curve)
1501:  <img src="rice_images/rice_image3_1.png" /> (Annex C - gel stages 1)
1505:  <img src="rice_images/rice_image3_2.png" /> (Annex C - gel stages 2)
1509:  <img src="rice_images/rice_image3_3.png" /> (Annex C - gel stages 3)
```

All are in `spec/examples/rice_images/` (directory confirmed to exist).

## Acceptance criteria

- All 5 images appear in `word/media/` in the DOCX
- Image relationships are in `word/_rels/document.xml.rels`
- Images render at correct dimensions in Word
- Local, relative-path images work; external URLs and data URIs are skipped gracefully
