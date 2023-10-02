# Convert DOCX

A nodejs utility that converts a docx file into multiple formats: pdf, html, text, md. Written as a proof-of-concept (POC) for converting Word documents in an AWS serverless architecture by adding node parsing libs in (a) lambda(s).

The output is not very good but each lib has options that can improve the quality.

## Parsing libraries

- [mammoth.js](https://github.com/mwilliamson/mammoth.js/) - docx to html
- [html-to-text](https://github.com/html-to-text/node-html-to-text)
- [puppeteer](https://github.com/puppeteer/puppeteer/tree/main) - html to pdf
- [turndown](https://github.com/mixmark-io/turndown) - html to md

## Usage

```bash
Usage: convert-docx [options]

CLI to convert docx files

Options:
  -f, --file <file>           add the specified file (default:
                              "test.docx")
  -t, --file-type <type>      type of file to convert: pdf, html, text, md
                              (default: "html")
  -o, --output-path <output>  output file path (default: "./")
  -V, --version               output the version number
  -h, --help                  display help for command
```
