package app;

import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Converter {

    private final static Logger LOGGER = Logger.getLogger(Converter.class.getName());

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
        try (InputStream in = new FileInputStream(new File(docIn)); XWPFDocument document = new XWPFDocument(in)) {
            out = new FileOutputStream(new File(docOut));
            PdfDocument pdfDocument;
            Document pdfDoc;

            PdfWriter writer = new PdfWriter(docOut);
            pdfDocument = new PdfDocument(writer);
            pdfDoc = new Document(pdfDocument);

            // TODO: Provide explanation
            int p = 0; // current paragraph
            int t = 0; // current table
            for (IBodyElement be : document.getBodyElements()) {
                if (be.getElementType() == BodyElementType.PARAGRAPH) {
                    pdfDoc.add(processParagraph(be.getBody().getParagraphArray(p)));
                    p++;
                } else if (be.getElementType() == BodyElementType.TABLE) {
                    pdfDoc.add(processTable(be.getBody().getTableArray(t)));
                    t++;
                }
            }

            pdfDocument.close();
            pdfDoc.close();
        }
        out.close();
    }

    /**
     * Processes the text in the given paragraph.
     *
     * @param par
     * @return iText paragraph
     */
    private Paragraph processParagraph(XWPFParagraph par) {
        Paragraph paragraph = new Paragraph();
        List<XWPFRun> runs = par.getRuns();
        runs.stream().map(run -> {
            Text text;

            // Text formatting
            text = formatText(run);

            // Text color
            text = colorText(run, text);

            // Adding images
            try {
                Image image = extractPictures(run);
                if (image != null) {
                    paragraph.add(image);
                }
            } catch (IOException ex) {
                LOGGER.log(Level.SEVERE, "Error while processing image", ex);
            }

            return text;
        }).forEachOrdered(text -> {
            paragraph.add(text);
        });
        return paragraph;
    }

    /**
     * Provides basic text formatting.
     *
     * @param run
     * @return formatted text
     */
    private Text formatText(XWPFRun run) {
        Text text = new Text(run.text());
        if (run.isBold()) {
            text.setBold();
        }
        if (run.isItalic()) {
            text.setItalic();
        }
        return text;
    }

    /**
     * Applies color to a text. Handles Exceptions, if no color is provided (e.g. auto).
     *
     * @param run
     * @param text
     * @return colored text
     */
    private Text colorText(XWPFRun run, Text text) {
        try {
            Color color = Helper.hexToRgb(run.getColor());
            text.setFontColor(color);
        } catch (NumberFormatException ex) {
            LOGGER.log(Level.FINE, ex.toString());
        } catch (Exception ex) {
            LOGGER.log(Level.FINE, "Color not found", ex);
        }
        return text;
    }

    /**
     * Extracts possible images.
     *
     * @param run
     * @return Extracted image
     * @throws IOException
     */
    private Image extractPictures(XWPFRun run) throws IOException {
        Image image = null;
        if (run.getEmbeddedPictures() != null) {
            for (XWPFPicture picture : run.getEmbeddedPictures()) {
                InputStream is = picture.getPictureData().getPackagePart().getInputStream();
                image = new Image(ImageDataFactory.create(is.readAllBytes()));

                // image settings
                image.setWidth((float) picture.getWidth());

                return image;
            }
        }
        return image;
    }

    private Table processTable(XWPFTable tab) throws IOException {
        List<Paragraph> paragraphs = new ArrayList<>();
        int columns = 0;
        int rows = 0;

        List<XWPFTableRow> rowsTemp = tab.getRows();
        for (XWPFTableRow rowTemp : rowsTemp) {
            rows++;
            List<XWPFTableCell> cellsTemp = rowTemp.getTableCells();
            for (XWPFTableCell cellTemp : cellsTemp) {
                columns++;
                paragraphs.add(new Paragraph(cellTemp.getText()));
            }
        }

        /*  A table needs to be read and built sequentially. 
        So the number of columns needs to be determined before creating a new table. */
        int numColumns = columns / rows;

        Table table = new Table(numColumns);
        paragraphs.forEach(p -> {
            table.addCell(p);
        });
        return table;
    }
}
