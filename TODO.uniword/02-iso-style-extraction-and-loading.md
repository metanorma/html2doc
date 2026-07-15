# Phase 2: ISO Style Extraction and Loading

## Goal
Extract the ISO 690:2021 DOCX fixture's complete style set (429 styles, numbering definitions, font table, theme) and make it loadable as a Uniword `StylesConfiguration` + `NumberingConfiguration` for use as the default DOCX output template.

## Background
The ISO fixture at `../uniword/spec/fixtures/uniword-private/fixtures/iso/ISO_690_2021-Word_document(en).docx` contains:
- **429 styles** in `word/styles.xml` (Finnish locale IDs, English names)
- **13 numbering definitions** in `word/numbering.xml` (decimal, bullet, annex numbering, heading numbering)
- **5 fonts** in `word/fontTable.xml` (Calibri, Cambria, Times New Roman, etc.)
- **Theme** in `word/theme/theme1.xml`
- **Settings** with mirror margins, even/odd headers, hyphenation

This represents the canonical ISO Word template. We need to extract and bundle these styles so every DOCX produced by html2doc uses this style set.

## Tasks

- [ ] Use Uniword to load the ISO fixture and extract styles:
  ```ruby
  pkg = Uniword::Ooxml::DocxPackage.from_file("ISO_690_2021-Word_document(en).docx")
  styles = pkg.styles          # StylesConfiguration with 429 styles
  numbering = pkg.numbering    # NumberingConfiguration with 13 definitions
  font_table = pkg.font_table  # FontTable
  theme = pkg.theme            # Theme
  ```

- [ ] Serialize the ISO styles to a YAML file: `config/styles/iso_styles.yml`
  - Use `StylesConfiguration.to_yaml` or manual extraction
  - This becomes the bundled default style source for DOCX output

- [ ] Serialize numbering to `config/styles/iso_numbering.yml`

- [ ] Create `Html2Doc::StyleLoader` module that:
  - Loads the ISO style set from YAML (or directly from a reference DOCX)
  - Returns a `Uniword::Wordprocessingml::StylesConfiguration`
  - Returns a `Uniword::Wordprocessingml::NumberingConfiguration`
  - Returns a `Uniword::Wordprocessingml::FontTable`
  - Returns a `Uniword::Drawingml::Theme`
  - Caches loaded instances (they're immutable templates)

- [ ] Map html2doc CSS class names to ISO OOXML style IDs:
  | html2doc CSS class | ISO OOXML styleId |
  |---|---|
  | `MsoNormal` / none | `Normaali` (Normal) |
  | `MsoTitle` | `zzSTDTitle` |
  | `MsoHeading1` | `Otsikko1` (Heading 1) |
  | `MsoHeading2` | `Otsikko2` |
  | `MsoHeading3` | `Otsikko3` |
  | `MsoTOC1`-`MsoTOC9` | `Sisluet1`-`Sisluet9` |
  | `MsoFootnoteText` | `Alaviitteenteksti` |
  | `MsoFootnoteReference` | `Alaviitteenviite` |
  | `MsoHeader` | `Yltunniste` |
  | `MsoFooter` | `Alatunniste` |
  | ISO-specific: `zzCover`, `zzCopyright`, etc. | Same styleId |
  | `Note`, `Example`, `Source`, etc. | Same styleId |

- [ ] Create a mapping configuration file: `config/styles/class_to_style_map.yml`

## Key Files to Create/Modify
- `config/styles/iso_styles.yml` — extracted ISO styles
- `config/styles/iso_numbering.yml` — extracted ISO numbering
- `config/styles/class_to_style_map.yml` — CSS class → OOXML style mapping
- `lib/html2doc/style_loader.rb` — style loading module
- `lib/html2doc/docx_converter.rb` — use StyleLoader

## Success Criteria
- ISO styles loadable as Uniword model objects
- Style lookup by CSS class name returns correct OOXML style ID
- No runtime errors loading the 429-style configuration
