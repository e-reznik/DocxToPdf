# Example usage
   
    String docIn = "myDoc.docx";
    String docOut = "myPdf.pdf";
    
    Converter c = new Converter();
    c.convert(docIn, docOut);

# Supported elements
- text
- images
- tables
  - ~~text formatting~~
- ~~headings~~
- ~~basic shapes~~

# Supported text formatting
- bold
- italic
- colors

# Dependencies

| Dependency | Version | Description                 |
|------------|---------|-----------------------------|
| Apache POI | 4.1.2   | Accessing Docx with XWPF |
| iText 7    | 7.1.12  | Generating PDF              |
