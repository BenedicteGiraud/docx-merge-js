# 📄 docx-merge

A fast and lightweight Node.js library written in TypeScript for merging two Microsoft Word (`.docx`) documents into one. Easily insert content at specific positions or based on placeholder patterns.

---

## 📦 Installation

Install via npm:

```bash
npm install @benedicte/docx-merge
```

📚 Dependencies

Only two lightweight dependencies:

- `adm-zip` – for extracting and rebuilding .docx (ZIP) files
- `fast-xml-parser` – for parsing and modifying the DOCX XML content

## 🛠️ API

```ts
mergeDocx(
  sourcePath: string,
  contentPath: string,
  options: {
    outputPath?: string;
    pattern?: string;
    insertStart?: boolean;
    insertEnd?: boolean;
  }
): void | buffer
```

Parameters:
- `sourcePath` (*equired*) – Path to the base `.docx` file
- `contentPath` (*equired*) – Path to the `.docx` file to insert into the base
- `options`:
    - `outputPath` – If provided, writes the merged document to this path. If omitted, returns a `Buffer`
    - `pattern` – String pattern in the source file to replace with the inserted content
    - `insertStart` – Insert the content at the **beginning** of the source file
    - `insertEnd` – Insert the content at the **end** of the source file

🔔 Note: You can combine pattern, insertStart, and insertEnd and at least one is required.

## 💡 Examples

### Replace a placeholder in the source DOCX

```ts
import { mergeDocx } from "@benedicte/docx-merge";

mergeDocx("./source.docx", "./table.docx", {
  outputPath: "./output.docx",
  pattern: "{{table}}",
});
```

### Get the merged result as a Buffer

```ts
import { mergeDocx } from "@benedicte/docx-merge";

const buffer = mergeDocx("./source.docx", "./table.docx", {
  pattern: "{{table}}",
});

// Use buffer (e.g., send as a response in a server)
```

### Insert at the start of the document

```ts
import { mergeDocx } from "@benedicte/docx-merge";

mergeDocx("./source.docx", "./table.docx", {
  outputPath: "./output.docx",
  insertStart: true,
});
```

### Insert at the end of the document

```ts
import { mergeDocx } from "@benedicte/docx-merge";

mergeDocx("./source.docx", "./table.docx", {
  outputPath: "./output.docx",
  insertEnd: true,
});
```

## 🧪 Testing

This project uses Vitest for unit testing.

To run tests:

```bash
npm run test
```

## 🔒 License

MIT

## 🤝 Contributing

Contributions, bug reports, and feature requests are welcome! Feel free to open an issue or submit a pull request.