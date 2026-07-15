require "html2doc"

html = <<-HTML
<html><head><title>Sample DOCX</title></head><body>
  <h1>Sample Document</h1>
  <h2>Introduction</h2>
  <p class="MsoNormal">This is a <b>bold</b> and <i>italic</i> paragraph with <u>underline</u> formatting.</p>
  <p class="MsoNormal">Second paragraph with <b><i>bold italic</i></b> text.</p>
  <h2>Table Example</h2>
  <table>
    <tr><th>Header 1</th><th>Header 2</th></tr>
    <tr><td>Cell 1</td><td>Cell 2</td></tr>
    <tr><td>Cell 3</td><td>Cell 4</td></tr>
  </table>
  <h2>List Example</h2>
  <p class="MsoListParagraphCxSpFirst" style="mso-list:l0 level1 lfo1">First item</p>
  <p class="MsoListParagraphCxSpMiddle" style="mso-list:l0 level1 lfo1">Second item</p>
  <p class="MsoListParagraphCxSpLast" style="mso-list:l0 level1 lfo1">Third item</p>
  <h2>Aligned Text</h2>
  <p class="MsoNormal" style="text-align:center">Centered text.</p>
  <p class="MsoNormal" style="text-align:right">Right-aligned text.</p>
  <p class="MsoNormal">Normal paragraph at the bottom.</p>
</body></html>
HTML

Html2Doc.new(
  filename: "/tmp/html2doc_sample",
  output_format: :docx,
).process(html)

puts "Generated /tmp/html2doc_sample.docx"

