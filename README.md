# Osysharp.Csv

Spreadsheets in and out of an Osy# app — the product list a shop owner edits in Excel, the subscriber list they export,
the stock count they paste back.

```osy
use Osysharp.Csv@0;          // in app.osy

var text  = WriteCsv(["SKU", "Name", "Price"], [["GG-1", "Kiambu AA, washed", "165"]]);
var table = ReadCsv(uploadedText);
foreach (var row in table.Rows) {
  var sku = table.Get(row, "SKU");            // by column name — case and spaces ignored
}
```

What it gets right so the file "just opens":

- **A UTF-8 byte-order mark**, so Excel reads `Göteborg` as `Göteborg`.
- **CRLF and RFC 4180 quoting** — a field is quoted only when it holds the separator, a quote or a line break.
- **The separator is detected** when reading: a Swedish or German Excel saves `;` (the comma is their decimal point),
  and a paste out of a spreadsheet is tab-separated.
- **No cell runs as a formula.** A value starting `=`, `+`, `-` or `@` is written with a leading `'` (CSV injection —
  `=HYPERLINK(…)` in a customer's name), and reading takes it off again. A number such as `-5` is left alone.
- **Row line numbers** (`row.Line`), so an import can say which line it refused.

## Examples

Taken from this package's own tests, which compile and run on every release.

```osy
using Osysharp.Csv;

[Test]
void a_file_excel_saved_is_read_field_for_field() {
  var text = "﻿SKU;Name;Price\r\nGG-1;\"Kiambu AA; washed\";165,00\r\nGG-2;\"The \"\"house\"\" blend\nline two\";149\r\n\r\n";
  var t = ReadCsv(text);
  Assert.Equal(";", t.Separator);
  Assert.Equal(3, t.Header.Count);
  Assert.Equal("SKU", t.Header[0]);
  Assert.Equal(2, t.Rows.Count);
  Assert.Equal("Kiambu AA; washed", t.Get(t.Rows[0], "name"));
  Assert.Equal("165,00", t.Get(t.Rows[0], "Price"));
  Assert.Equal("The \"house\" blend\nline two", t.Get(t.Rows[1], "Name"));
  Assert.Equal(2, t.Rows[0].Line);
  Assert.Equal(3, t.Rows[1].Line);
  Assert.Equal("", t.Get(t.Rows[0], "Colour"));
}
```

From [`tests/csv.test.osy`](tests/csv.test.osy).

```osy
using Osysharp.Csv;

[Test]
void what_is_written_reads_back_the_same() {
  List<List<string>> rows = [["GG-1", "Kiambu AA, washed", "165"], ["GG-2", "Say \"hi\"", "149"], ["GG-3", "two\nlines", ""]];
  var text = WriteCsv(["SKU", "Name", "Price"], rows);
  Assert.True(text.StartsWith("﻿SKU,Name,Price\r\n"));
  Assert.Contains("GG-1,\"Kiambu AA, washed\",165\r\n", text);
  var t = ReadCsv(text);
  Assert.Equal(",", t.Separator);
  Assert.Equal(3, t.Rows.Count);
  for (var r = 0; r < 3; r++) {
    for (var f = 0; f < 3; f++) { Assert.Equal(rows[r][f], t.Rows[r].Fields[f]); }
  }
}
```

From [`tests/csv.test.osy`](tests/csv.test.osy).

```osy
using Osysharp.Csv;

[Test]
void a_formula_is_written_as_text_and_a_number_is_not() {
  var text = WriteCsv(["Name", "Delta"], [["=HYPERLINK(\"http://evil.example\",\"Click\")", "-5"], ["@SUM(A1)", "+3.5"], ["- on sale", "0"]]);
  Assert.Contains("\"'=HYPERLINK(\"\"http://evil.example\"\",\"\"Click\"\")\",-5", text);
  Assert.Contains("'@SUM(A1),+3.5", text);
  Assert.Contains("'- on sale,0", text);
  var t = ReadCsv(text);
  Assert.Equal("=HYPERLINK(\"http://evil.example\",\"Click\")", t.Rows[0].Fields[0]);
  Assert.Equal("-5", t.Rows[0].Fields[1]);
  Assert.Equal("@SUM(A1)", t.Rows[1].Fields[0]);
  Assert.Equal("- on sale", t.Rows[2].Fields[0]);
}
```

From [`tests/csv.test.osy`](tests/csv.test.osy).

## Learn more

- [Osysharp.Csv in the package catalogue](https://osysharp.com/packages/#osysharp-csv)
- [The Osy# reference](https://osysharp.com/reference/) — the language, the runtime and every kit
- [`llms-full.txt`](https://osysharp.com/llms-full.txt) — the whole reference and specification as one Markdown file, for a coding agent

## Licence

MIT — see [`LICENSE`](LICENSE).
