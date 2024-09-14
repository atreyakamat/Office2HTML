import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class XlsConverter {
    public void convert(String inputPath, File baseFolder) throws IOException {
        // Create the HTML output file inside the base folder
        File htmlFile = new File(baseFolder, baseFolder.getName() + ".html");

        try (FileInputStream fis = new FileInputStream(inputPath);
             HSSFWorkbook workbook = new HSSFWorkbook(fis);
             PrintWriter writer = new PrintWriter(new FileOutputStream(htmlFile))) {

            // Add Bootstrap CSS and other HTML structure
            writer.println("<!DOCTYPE html>");
            writer.println("<html lang=\"en\">");
            writer.println("<head>");
            writer.println("<meta charset=\"UTF-8\">");
            writer.println("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            writer.println("<title>Excel to HTML</title>");
            writer.println("<link rel=\"stylesheet\" href=\"https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css\">");
            writer.println("<style>");
            writer.println("table, td, th { border: 1px solid white; border-collapse: collapse; }");
            writer.println("td, th { padding: 8px; background-color: #FFFFFF; text-align: center; }");
            writer.println("body { text-align: left; }"); // Align body content to the left
            writer.println("</style>");
            writer.println("<script>");
            writer.println("    function functionOne() {console.log(\"Hello\");}");
            writer.println("    function functionTwo() {console.log(\"Hello\");}");
            writer.println("    function functionThree() {console.log(\"Hello\");}");
            writer.println("    function functionFour() {console.log(\"Hello\");}");
            writer.println("    function functionFive() {console.log(\"Hello\");}");
            writer.println("</script>");
            writer.println("</head>");
            writer.println("<body>");
            writer.println("<div class=\"my-4\">"); // Removed 'container' to prevent centering
            writer.println("<h2 class=\"text-left mb-4\">Converted Excel Data</h2>"); // Aligned header to the left
            writer.println("<table class=\"table table-bordered table-hover\">");

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                writer.println("<thead class=\"thead-dark\"><tr>");
                Row headerRow = sheet.getRow(0); // Assuming the first row is the header

                // Write header row if available
                if (headerRow != null) {
                    for (Cell headerCell : headerRow) {
                        writer.print("<th>");
                        writer.print(headerCell.toString());
                        writer.println("</th>");
                    }
                    writer.println("</tr></thead>");
                }

                writer.println("<tbody>");
                // Start from row 1 if header is present
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) continue;

                    writer.println("<tr>");
                    for (Cell cell : row) {
                        writer.print("<td>");
                        switch (cell.getCellType()) {
                            case STRING:
                                writer.print(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                writer.print(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                writer.print(cell.getBooleanCellValue());
                                break;
                            default:
                                writer.print(" ");
                        }
                        writer.print("</td>");
                    }
                    writer.println("</tr>");
                }
                writer.println("</tbody>");
            }

            writer.println("</table>");
            writer.println("</div>"); // End container div
            writer.println("</body>");
            writer.println("</html>");
        }
    }
}
