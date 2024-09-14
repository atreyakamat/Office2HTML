import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

public class DocConverter {

    public void convert(String inputPath, File baseFolder) throws IOException {
        // Create the HTML output file inside the base folder
        File htmlFile = new File(baseFolder, baseFolder.getName() + ".html");

        // Folder to store extracted images inside the base folder
        File imageFolder = new File(baseFolder, "images");
        if (!imageFolder.exists()) {
            boolean imageFolderCreated = imageFolder.mkdir();
            if (!imageFolderCreated) {
                throw new IOException("Error: Could not create image folder.");
            }
        }

        try (FileInputStream fis = new FileInputStream(inputPath);
             HWPFDocument doc = new HWPFDocument(fis);
             PrintWriter writer = new PrintWriter(new FileOutputStream(htmlFile))) {

            // Extract the document's text and table content
            Range range = doc.getRange();
            TableIterator tableIterator = new TableIterator(range);

            // Extract pictures from the document
            PicturesTable picturesTable = doc.getPicturesTable();
            List<Picture> pictures = picturesTable.getAllPictures();
            int imageIndex = 0;

            // Begin writing HTML content
            writer.println("<html>");
            writer.println("<head><style>body { font-family: Arial, sans-serif; }</style>");
            writer.println("<script>");
            writer.println("    function functionOne() {console.log(\"Hello\");}");
            writer.println("    function functionTwo() {console.log(\"Hello\");}");
            writer.println("    function functionThree() {console.log(\"Hello\");}");
            writer.println("    function functionFour() {console.log(\"Hello\");}");
            writer.println("    function functionFive() {console.log(\"Hello\");}");
            writer.println("</script>");
            writer.println("</head>");
            writer.println("<body>");

            // Iterate through paragraphs and tables
            int currentTableIndex = 0;
            for (int i = 0; i < range.numParagraphs(); i++) {
                // If there's a table at this position, handle it
                if (tableIterator.hasNext()) {
                    Table table = tableIterator.next();
                    if (range.getParagraph(i).isInTable()) {
                        // Start table in HTML
                        writer.println("<table border='1' style='border-collapse: collapse;'>");

                        // Loop through table rows
                        for (int rowIndex = 0; rowIndex < table.numRows(); rowIndex++) {
                            TableRow row = table.getRow(rowIndex);
                            writer.println("<tr>");

                            // Loop through table cells
                            for (int cellIndex = 0; cellIndex < row.numCells(); cellIndex++) {
                                TableCell cell = row.getCell(cellIndex);
                                writer.print("<td>");
                                writer.print(cell.text().trim());  // Add cell text
                                writer.println("</td>");
                            }
                            writer.println("</tr>");
                        }
                        writer.println("</table>");
                    }
                }

                // Handle normal paragraphs
                if (!range.getParagraph(i).isInTable()) {
                    String paragraphText = range.getParagraph(i).text();
                    writer.println("<p>" + paragraphText.trim() + "</p>");
                }
            }

            // Handle images
            for (Picture picture : pictures) {
                String imageFileName = baseFolder.getName() + "-image-" + (imageIndex++) + "." + picture.suggestFileExtension();
                File imageFile = new File(imageFolder, imageFileName);

                // Write image to the image folder
                try (FileOutputStream imageOutputStream = new FileOutputStream(imageFile)) {
                    imageOutputStream.write(picture.getContent());
                }

                // Insert image reference in HTML
                writer.println("<p><img src='" + imageFolder.getName() + "/" + imageFileName + "' alt='Image' /></p>");
            }

            // End HTML content
            writer.println("</body>");
            writer.println("</html>");
        }
    }
}
