# Example usage
   
    String docIn = "myDoc.docx";
    String docOut = "myPdf.pdf";
    
    Converter c = new Converter();
    c.convert(docIn, docOut);

# Supported Elements
- text
- images
- ~~tables~~
- ~~headings~~
- ~~basic shapes~~

# Supported Formatting
- bold
- italic
- colors

# Dependencies

| Dependency | Version | Description                 |
|------------|---------|-----------------------------|
| Apache POI | 4.1.2   | Accessing Docx with XWPF |
| iText 7    | 7.1.12  | Generating PDF              |
