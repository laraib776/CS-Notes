# API Mastery Course Notes Generator

This project uses Node.js to generate a polished Microsoft Word document (`.docx`) from JavaScript code. The main script, `api_course.js`, builds the document structure with the [`docx`](https://www.npmjs.com/package/docx) package and writes the final file to the project folder.

## What This Project Does

- Creates a complete DOCX course-notes document.
- Uses headings, paragraphs, tables, colors, page layout, headers, footers, and styled text.
- Generates the output file automatically when the script is run.
- Keeps the document content and layout controlled from JavaScript.

Generated file:

```text
API_Mastery_Complete_Course_Notes.docx
```

## Requirements

Before running the project, make sure these are installed:

- Node.js
- npm
- A code editor such as VS Code

Check your versions:

```bash
node -v
npm -v
```

## Install Dependencies

From the project folder, run:

```bash
npm install
```

This installs the required package:

```text
docx
```

The `docx` package is used to create Word documents directly from JavaScript.

## How To Convert The JS File Into A DOCX File

Run this command:

```bash
node api_course.js
```

If everything is working, the terminal should show something like:

```text
Done! Wrote C:\Users\AB\Desktop\Projects\NOTES\API_Mastery_Complete_Course_Notes.docx
```

After that, open the generated `.docx` file in Microsoft Word, Google Docs, LibreOffice, or any other Word-compatible editor.

## How The Script Works

The script follows this general process:

1. Imports the required tools from the `docx` package.
2. Defines document colors, borders, styles, and helper functions.
3. Creates headings, paragraphs, tables, and other document blocks.
4. Builds a complete `Document` object.
5. Converts that document into a file buffer using `Packer.toBuffer`.
6. Saves the buffer as a `.docx` file using Node's `fs.writeFileSync`.

The final save logic looks like this:

```js
Packer.toBuffer(doc).then(buffer => {
  const outputPath = path.join(__dirname, 'API_Mastery_Complete_Course_Notes.docx');
  fs.writeFileSync(outputPath, buffer);
  console.log(`Done! Wrote ${outputPath}`);
}).catch(console.error);
```

Using `path.join(__dirname, ...)` makes the output path work correctly on Windows, macOS, and Linux.

## Common Error: ENOENT

If you see an error like this:

```text
Error: ENOENT: no such file or directory
```

It usually means the script is trying to save the file in a folder that does not exist.

Example of a problematic path on Windows:

```text
/mnt/user-data/outputs/file.docx
```

Fix it by saving the file inside the current project folder:

```js
const outputPath = path.join(__dirname, 'output.docx');
```

## Project Structure

```text
NOTES/
├── api_course.js
├── package.json
├── package-lock.json
├── README.md
└── API_Mastery_Complete_Course_Notes.docx
```

For GitHub, you usually commit the source files and avoid committing `node_modules`.

Recommended files to commit:

```text
api_course.js
package.json
package-lock.json
README.md
```

Usually do not commit:

```text
node_modules/
```

Committing the generated `.docx` file is optional. Commit it if you want people to download the final document directly from the repository.

## Converting Other File Types Into More Readable Files

The same idea can be used to convert many file types into cleaner, more readable formats.

### Text Or Markdown To DOCX

Use this when you have notes in `.txt` or `.md` format and want a Word document.

Possible approach:

- Read the source file with `fs.readFileSync`.
- Split content into headings, paragraphs, lists, or code blocks.
- Convert each part into `docx` paragraphs.
- Save the result as `.docx`.

Useful packages:

```text
docx
marked
markdown-it
```

### JSON To DOCX Or HTML

Use this when your data is structured, such as API responses, course modules, user records, or documentation content.

Possible readable outputs:

- DOCX report
- HTML page
- Markdown documentation
- PDF report

Useful packages:

```text
docx
ejs
handlebars
puppeteer
```

### CSV Or Excel To Reports

Use this when you have spreadsheet-like data and want readable summaries.

Possible outputs:

- Tables in DOCX
- Summary reports
- Charts in HTML or PDF

Useful packages:

```text
csv-parse
xlsx
docx
```

### HTML To PDF

Use this when you already have a designed webpage or report and want a printable file.

Useful package:

```text
puppeteer
```

Basic flow:

1. Create an HTML page.
2. Style it with CSS.
3. Use Puppeteer to export it as PDF.

### Markdown To HTML

Use this when you want notes or documentation to become a readable web page.

Useful packages:

```text
marked
markdown-it
```

Basic flow:

1. Read the Markdown file.
2. Convert Markdown to HTML.
3. Save the result as an `.html` file.

## Best Practices For File Conversion Projects

- Keep source content separate from output files when possible.
- Use clear output filenames.
- Use `path.join` instead of hardcoded system paths.
- Create output folders before writing files if needed.
- Add helpful terminal messages after files are created.
- Use templates when converting repeated content.
- Add a README so other people know how to install and run the project.
- Avoid committing `node_modules` to GitHub.

## Suggested `.gitignore`

Create a `.gitignore` file before pushing to GitHub:

```gitignore
node_modules/
.env
*.log
```

If you do not want to commit generated Word files, also add:

```gitignore
*.docx
```

## Useful Commands

Install dependencies:

```bash
npm install
```

Generate the DOCX file:

```bash
node api_course.js
```

Check Git status:

```bash
git status
```

Add files to Git:

```bash
git add api_course.js package.json package-lock.json README.md
```

Commit changes:

```bash
git commit -m "Add DOCX course notes generator"
```

## Notes For Future Improvements

- Move course content into a separate Markdown, JSON, or text file.
- Keep styling helpers in a separate file.
- Add multiple output formats, such as DOCX, HTML, and PDF.
- Add a command in `package.json`, such as `npm run build`.
- Add an `outputs/` folder for generated files.

