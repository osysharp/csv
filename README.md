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

