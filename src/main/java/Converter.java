
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Text;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

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
                List<XWPFRun> runs = p.getRuns();
                Paragraph par = new Paragraph();
                /* Each chunk (word) can have its own formatting. Therefore each of them will be iterated. */
                runs.stream().map(run -> {
                    // Text formatting
                    Text text = formatText(run);
                    // Text color
                    text = colorText(run, text);
                    return text;
                }).forEachOrdered(text -> {
                    par.add(text);
                });
                pdfDoc.add(par);
            });

            pdfDocument.close();
            pdfDoc.close();
        }
        out.close();
    }

    /**
     * Provides basic text formatting.
     *
     * @param run
     * @return formatted text
     */
    private Text formatText(XWPFRun run) {
        Text text = new Text(run.text());
        System.out.println(text);
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
            Color color = hexToRgb(run.getColor());
            text.setFontColor(color);
        } catch (NumberFormatException ex) {
            LOGGER.log(Level.FINE, ex.toString());
        } catch (Exception ex) {
            LOGGER.log(Level.FINE, "Color not found", ex);
        }
        return text;
    }

    /**
     * Converts HEX to RGB
     *
     * @param hex
     * @return
     * @throws NumberFormatException if no color is provided (e.g. auto)
     * @throws Exception if an unexpected exception occurs
     */
    private Color hexToRgb(String hex) throws NumberFormatException, Exception {
        if (hex != null && hex.length() != 6) {
            throw new NumberFormatException(hex + " is not a color");
        }
        int r = Integer.valueOf(hex.substring(0, 2), 16);
        int g = Integer.valueOf(hex.substring(2, 4), 16);
        int b = Integer.valueOf(hex.substring(4, 6), 16);

        Color rgb = new DeviceRgb(r, g, b);
        return rgb;
    }
}
