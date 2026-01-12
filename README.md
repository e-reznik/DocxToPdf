# DocxToPdf

> **Note:** This project is archived and no longer maintained. A newer, improved version is available at [Docx2PDF](https://github.com/e-reznik/Docx2PDF), which uses [DocxJavaMapper](https://github.com/e-reznik/DocxJavaMapper) for DOCX parsing instead of Apache POI.

A simple Java library that converts Microsoft Word documents (.docx) to PDF format using Apache POI and iText 7.

## Usage

```java
String docIn = "myDoc.docx";
String docOut = "myPdf.pdf";

Converter c = new Converter();
c.convert(docIn, docOut);
```

## Supported Features

| Feature | Status |
|---------|--------|
| Text | Supported |
| Bold / Italic | Supported |
| Text colors | Supported |
| Images | Supported |
| Tables | Supported (text only) |
| Table text formatting | Not supported |
| Headings | Not supported |
| Basic shapes | Not supported |

## Dependencies

| Dependency | Version | Description |
|------------|---------|-------------|
| Apache POI | 4.1.2 | Reading DOCX files via XWPF |
| iText 7 | 7.1.12 | Generating PDF output |

## Successor Project

This project has been superseded by [Docx2PDF](https://github.com/e-reznik/Docx2PDF), which offers:

- No dependency on Apache POI
- Uses [DocxJavaMapper](https://github.com/e-reznik/DocxJavaMapper) for mapping DOCX structure to Java objects
- Cleaner architecture with separation between parsing and PDF generation

## License

MIT
