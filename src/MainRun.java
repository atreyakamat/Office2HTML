import java.io.File;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

public class MainRun {

    public static void main(String[] args) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select a file to convert");
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Document files (.doc, .docx, .xls, .xlsx)", "doc", "docx", "xls", "xlsx");
        fileChooser.setFileFilter(filter);
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            String inputPath = selectedFile.getAbsolutePath();

            // Create a base folder for output in the same directory as the input file
            String baseFolderPath = inputPath.replaceAll("\\.(doc|docx|xls|xlsx)$", "");
            File baseFolder = new File(baseFolderPath);
            if (!baseFolder.exists()) {
                boolean folderCreated = baseFolder.mkdir();
                if (!folderCreated) {
                    System.out.println("Error: Could not create base folder.");
                    return;
                }
            }

            // Define the output paths for the HTML file and images folder
            String outputPath = baseFolderPath + ".html";

            try {
                String fileType = getFileExtension(inputPath);

                switch (fileType) {
                    case "xls" -> {
                        XlsConverter xlsConverter = new XlsConverter();
                        xlsConverter.convert(inputPath, baseFolder);
                        System.out.println(".xls file successfully converted to HTML.");
                    }
                    case "xlsx" -> {
                        XlsxConverter xlsxConverter = new XlsxConverter();
                        xlsxConverter.convert(inputPath, baseFolder);
                        System.out.println(".xlsx file successfully converted to HTML.");
                    }
                    case "doc" -> {
                        DocConverter docConverter = new DocConverter();
                        docConverter.convert(inputPath, baseFolder);
                        System.out.println(".doc file successfully converted to HTML.");
                    }
                    case "docx" -> {
                        DocxConverter docxConverter = new DocxConverter();
                        docxConverter.convert(inputPath, baseFolder);
                        System.out.println(".docx file successfully converted to HTML.");
                    }
                    default -> System.out.println("Unsupported file format. Please input .xls, .xlsx, .doc, or .docx.");
                }
            } catch (IOException e) {
                System.out.println("Error during conversion: " + e.getMessage());
            }
        } else {
            System.out.println("No file selected.");
        }
    }

    private static String getFileExtension(String fileName) {
        int lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex != -1 && lastDotIndex < fileName.length() - 1) {
            return fileName.substring(lastDotIndex + 1).toLowerCase();
        } else {
            return "";
        }
    }
}
