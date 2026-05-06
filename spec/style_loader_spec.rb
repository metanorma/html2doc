require "spec_helper"
require "html2doc"

RSpec.describe Html2Doc::StyleLoader do
  describe ".iso_styles" do
    it "loads ISO styles without errors" do
      styles = described_class.iso_styles
      expect(styles).not_to be_nil
    end

    it "returns the same object on repeated calls (cached)" do
      first = described_class.iso_styles
      second = described_class.iso_styles
      expect(first).to equal(second)
    end
  end

  describe ".iso_numbering" do
    it "loads ISO numbering without errors" do
      numbering = described_class.iso_numbering
      expect(numbering).not_to be_nil
    end
  end

  describe ".iso_font_table" do
    it "loads ISO font table without errors" do
      ft = described_class.iso_font_table
      expect(ft).not_to be_nil
    end
  end

  describe ".iso_settings" do
    it "loads ISO settings without errors" do
      settings = described_class.iso_settings
      expect(settings).not_to be_nil
    end
  end

  describe ".iso_theme" do
    it "loads ISO theme without errors" do
      theme = described_class.iso_theme
      expect(theme).not_to be_nil
    end
  end

  describe ".style_for_class" do
    it "maps MsoNormal to Normal" do
      expect(described_class.style_for_class("MsoNormal")).to eq("Normal")
    end

    it "maps MsoTitle to zzSTDTitle" do
      expect(described_class.style_for_class("MsoTitle")).to eq("zzSTDTitle")
    end

    it "maps MsoHeading1 to Heading1" do
      expect(described_class.style_for_class("MsoHeading1")).to eq("Heading1")
    end

    it "maps MsoFootnoteText to FootnoteText" do
      expect(described_class.style_for_class("MsoFootnoteText")).to eq("FootnoteText")
    end

    it "maps Hyperlink to Hyperlink" do
      expect(described_class.style_for_class("Hyperlink")).to eq("Hyperlink")
    end

    it "maps ISO-specific classes" do
      expect(described_class.style_for_class("zzCover")).to eq("zzCover")
      expect(described_class.style_for_class("BiblioTitle")).to eq("BiblioTitle")
      expect(described_class.style_for_class("Terms")).to eq("Terms")
    end

    it "returns nil for unknown classes" do
      expect(described_class.style_for_class("UnknownClass")).to be_nil
    end
  end

  describe ".class_to_style" do
    it "returns a hash" do
      expect(described_class.class_to_style).to be_a(Hash)
    end

    it "is frozen" do
      expect(described_class.class_to_style).to be_frozen
    end

    it "has entries for standard Mso classes" do
      map = described_class.class_to_style
      expect(map).to have_key("MsoNormal")
      expect(map).to have_key("MsoHeading1")
      expect(map).to have_key("MsoTOC1")
    end
  end
end
