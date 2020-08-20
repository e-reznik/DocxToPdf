# DocxToPdf
A very basic Docx to PDF converter. Currently supports only simple text.

# Examples usage

    Converter c = new Converter();
    String docIn = "/home/user/myDoc.docx";
    String docOut = "/home/user/myPdf.pdf";
   
    c.convert(docIn, docOut);
