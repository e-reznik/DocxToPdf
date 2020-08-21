# DocxToPdf
A very basic Docx to PDF converter. Currently supports only simple text.

# Examples usage

    Converter c = new Converter();
    String docIn = "myDoc.docx";
    String docOut = "myPdf.pdf";
   
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
