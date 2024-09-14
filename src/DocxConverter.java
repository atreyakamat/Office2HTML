import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

public class DocxConverter {

    public void convert(String inputPath, File baseFolder) throws IOException {
        // Create the HTML output file inside the base folder
        File htmlFile = new File(baseFolder, baseFolder.getName() + ".html");

        // File name without extension for image folder creation
        String baseName = new File(inputPath).getName().replaceFirst("[.][^.]+$", "");

        // Folder to store extracted images
        File imageFolder = new File(baseFolder, baseName + "-images");
        if (!imageFolder.exists()) {
            imageFolder.mkdir();
        }
        
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument doc = new XWPFDocument(fis);
             PrintWriter writer = new PrintWriter(new FileOutputStream(htmlFile))) {
            JavaScriptImplementation jsImpl = new JavaScriptImplementation();
            // Begin writing HTML content
            writer.println("<html>");
            writer.println("<head><style>body { font-family: Arial, sans-serif; }</style>");
            jsImpl.writeScript(writer); 
            writer.println("</head>");
            writer.println("<body>");

            // Loop through body elements (paragraphs, tables) in sequence
            List<IBodyElement> bodyElements = doc.getBodyElements();
            for (IBodyElement element : bodyElements) {
                switch (element.getElementType()) {
                    case PARAGRAPH:
                        XWPFParagraph paragraph = (XWPFParagraph) element;
                        extractImages(paragraph, doc, imageFolder, baseName, writer); // Extract any images in the paragraph
                        writer.println("<p>" + paragraph.getText() + "</p>");  // Convert paragraph text to HTML
                        break;
                    case TABLE:
                        XWPFTable table = (XWPFTable) element;
                        writer.println("<table border='1'>");
                        for (XWPFTableRow row : table.getRows()) {
                            writer.println("<tr>");
                            for (XWPFTableCell cell : row.getTableCells()) {
                                String cellText = cell.getText();
                                int colSpan = 1;

                                // Check for cell properties like gridSpan
                                CTTcPr tcPr = cell.getCTTc().getTcPr();
                                if (tcPr != null && tcPr.isSetGridSpan()) {
                                    colSpan = tcPr.getGridSpan().getVal().intValue();
                                }

                                // Write the table cell in HTML, handling colSpan
                                writer.printf("<td colspan='%d'>%s</td>%n", colSpan, cellText);
                            }
                            writer.println("</tr>");
                        }
                        writer.println("</table>");
                        break;
                    default:
                        // Handle other types of elements (if needed)
                        break;
                }
            }

            // End HTML content
            writer.println("</body>");
            writer.println("</html>");
        }
    }

    // Method to extract images from paragraphs and convert them to HTML image tags
    private void extractImages(XWPFParagraph paragraph, XWPFDocument doc, File imageFolder, String baseName, PrintWriter writer) throws IOException {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs != null) {
            int imageIndex = 0;
            for (XWPFRun run : runs) {
                List<XWPFPicture> pictures = run.getEmbeddedPictures();
                for (XWPFPicture picture : pictures) {
                    XWPFPictureData pictureData = picture.getPictureData();
                    String imageFileName = baseName + "-image-" + (imageIndex++) + "." + pictureData.suggestFileExtension();
                    File imageFile = new File(imageFolder, imageFileName);

                    // Write image to the image folder
                    try (FileOutputStream imageOutputStream = new FileOutputStream(imageFile)) {
                        imageOutputStream.write(pictureData.getData());
                    }

                    // Insert image reference in HTML
                    writer.println("<p><img src='" + imageFolder.getName() + "/" + imageFileName + "' alt='Image' /></p>");
                }
            }
        }
    }
}
