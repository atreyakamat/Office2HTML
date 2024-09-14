import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxConverter {
    public void convert(String inputPath, File baseFolder) throws IOException {
        // Create the HTML output file inside the base folder
        File htmlFile = new File(baseFolder, baseFolder.getName() + ".html");

        try (FileInputStream fis = new FileInputStream(inputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             PrintWriter writer = new PrintWriter(new FileOutputStream(htmlFile))) {
            JavaScriptImplementation jsImpl = new JavaScriptImplementation();

            // Adding Bootstrap CSS link to the HTML
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
            jsImpl.writeScript(writer); 
            writer.println("</head><body>");
            writer.println("<div class='container mt-5'>");
            writer.println("<table class='table table-bordered table-striped'>");  // Using Bootstrap table classes

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                for (Row row : sheet) {
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
            }
            writer.println("</table>");
            writer.println("</div></body></html>");
        }
    }
}
