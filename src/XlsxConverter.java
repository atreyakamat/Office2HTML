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

            // Adding Bootstrap CSS link to the HTML
            writer.println("<html><head>");
            writer.println("<link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css\" rel=\"stylesheet\">");
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
