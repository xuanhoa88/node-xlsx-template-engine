const fs = require("fs");
const path = require("path");
const XlsxTemplate = require("./main");

const ORIGIN_DIR = path.resolve(__dirname, "../");
const TEMPLATE_FILE = path.join(
  ORIGIN_DIR,
  "test-data",
  "test-named-tables.xlsx"
);

// Create a template
const template = new XlsxTemplate();
template.loadTemplate(TEMPLATE_FILE);

// Replacements take place on first sheet
const sheetNumber = 1;

// Set up some placeholder values matching the placeholders in the template
const values = {
  extractDate: new Date(),
  dates: [
    new Date("2013-06-01"),
    new Date("2013-06-02"),
    new Date("2013-06-03"),
  ],
  people: [
    {
      name: "John Smith",
      age: 20,
    },
    { name: "Bob Johnson", age: 22 },
  ],
  sales: [{ payments: [123, 456] }],

  ages: [
    { name: "John", age: 10 },
    {
      name: "Bill",
      age: 12,
      avatar:
        "iVBORw0KGgoAAAANSUhEUgAAALAAAAA2CAYAAABnXhObAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAUjSURBVHhe7ZtbyGVjGMfXlmmccpqRcYgZmQsTMeVYQiJzozEYJRHCcCUaNy5kyg3KhUNEhjSRbwipIcoF0pRpxiFiLsYUIqcoZ8b2/6/1vKtnv3u9+/smuXjW/v/qv593PetZa+/17f9617vetb9BVVVDSIiQ7GFRiJC0PfBwOGRbiBAMBoPat+qBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoZGBRWhkYBEaGViERgYWoenND9oHg8EFCOdBPJ57cDzfWf5MhLPZBr9AL2LdzmZxFNTui3AdtF+dqKqtqH3F2iOg9g6EVLcFdTPW7gT1CxAuhE6FFkG/Ql9Cm7HtC4gtqL0S4ehmaRzU32XNiWA/VyGcDGGT4c11sifg2Np/hWODB8iXsAJ3p2OBznL5rS5Pfey38wJPuzpqe6HuXFdDfdRVlwRo9r+stkvboYWu/ifLlzTP7z8XOBZ63mqTDuyqjap0XNMwhGCvStjjkeNw9q6ydgtyCxEub5Za5lnMWW0xsQzbn2btEZBfh0DtWSeq6h3oQeghaBMTYCl0UtMc4U/obegN6FXoJWgGXyBPhk7wfrcgfAJdXCemgNrJ3t0RBUo98KeW+wCiIdjmMCLfnpdYNrw+y+us9ltb/6FFisOWvO4Ut/4HaHWh5j7odJdLPfAXvnYuAk/Ztuuhn61NqQcOzt/QM02zWome6gBrJzj2Jey9imC7lQjsrcn1FsmlFj33WiS34Q+/0dotyL0L3QptttR/hT32cuzvWkQec6+ZNgNvaJo1ay3SlLyxOr5Zqm6wWKI1qpluW7NULcF+zrE29zkfId08sid/3Nq7BfazwMvSRfA+j0Lv2eIui71lqgyML/Z1xO+bxepqi+RGizTAW9Ycw0yZxr+8VJP7LZJLLJJjLBIOY2qwj1XQRtOM6VmIMxM5R0CcTWmFunRSzAX1wD0ifZnpMn4kzHCita+xyLEoKX3xNC9NTB7mCwz/ZL3U4A18sEWSThpyCMRenOL+qMsgfxNXD/IKTFqXox64RyRTcqossQ4mvsna5AmLacYinxv3sw9/YNvlFNpp/HoYlldY+2uLhLMMiTchvifFGYlJcJ6YJ0IrnDDcfq703sCkvpsDfAkrMNssxGsut8Ny1O8W2xkHQPOxsdPleOPGxmx6zG3Dyz4bfI9FKe/WP2DrqRUu/6PldnsWwgvwgU3av2YhguOHBX42IA0J/Fj2N4se3/tOwg8j0qwH3+POpjnC/z1G1RCipzxn0eMfBU8y8C70AINcyKeT4iAMIy6y9u3QN02zWoP8y9D50OHQMuTOaFaVQd0JuWxVJ1i/N7SUwqI/QY5CbjG0vy33hrorBmPddCSB2YYQm7L69y3ftW6L5eshBFhiy9R6X5sEOD+caja4PG8QeROX1pXUNYQoaX6qzQV4JcnrvTo/fzSl45nWHpj4YUT+Q5y8B+YPhRJ8nDsG/qjMf9UsVVdYZJ43hpxS4+NjPsHL4Xj8EdTxUXEi7adE/Q0WKD5mNno1rOClr7FzcxkMDS6P+zDiWNIsQg3yeyHHG6kRmGcsrfP50r5zZqvD+sUIh0L8ZdznqONj4zHSZ+ui6/N6sG2xY8K2/1gzNDjG2re9MrCYHpKBp3kIIXqADCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAiNDCxCIwOL0MjAIjQysAhN+x8ZQkREPbAITFX9C5ozpqaetbGcAAAAAElFTkSuQmCC",
    },
  ],
  days: ["Monday", "Tuesday", "Wednesday"],
  hours: [
    { name: "Bob", days: [10, 20, 30] },
    { name: "Jim", days: [12, 24, 36] },
  ],
  progress: 100,
};

// Perform substitution
template.substitute(sheetNumber, values);

fs.writeFileSync(
  path.join(ORIGIN_DIR, "test-data/output", path.basename(TEMPLATE_FILE)),
  // Get binary data
  template.generate(),
  "binary"
);
