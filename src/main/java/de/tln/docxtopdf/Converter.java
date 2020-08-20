package de.tln.docxtopdf;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class Converter {

    public Converter() {
    }

    /**
     * Converts Docx into PDF.
     *
     * @param docIn Path to the Docx
     * @param docOut Path, where the PDF will be saved
     * @throws FileNotFoundException
     * @throws IOException
     */
    public void convert(String docIn, String docOut) throws FileNotFoundException, IOException {
        OutputStream out;
        try (InputStream in = new FileInputStream(new File(docIn));
                XWPFDocument document = new XWPFDocument(in)) {
            out = new FileOutputStream(new File(docOut));
            PdfDocument pdfDocument;
            Document pdfDoc;

            PdfWriter writer = new PdfWriter(docOut);
            pdfDocument = new PdfDocument(writer);
            pdfDoc = new Document(pdfDocument);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            paragraphs.forEach(p -> {
                pdfDoc.add(new Paragraph(p.getText()));
            });

            pdfDocument.close();
            pdfDoc.close();
        }
        out.close();
    }
}
