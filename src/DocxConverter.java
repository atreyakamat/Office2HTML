import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

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

            // Begin writing HTML content
            writer.println("<html>");
            writer.println("<head><style>body { font-family: Arial, sans-serif; }</style></head>");
            writer.println("<body>");

            // Extract paragraphs and images
            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            List<XWPFPictureData> pictures = doc.getAllPictures();
            int imageIndex = 0;

            // Loop through paragraphs and convert them to HTML paragraphs
            for (XWPFParagraph paragraph : paragraphs) {
                writer.println("<p>");
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.toString();
                    writer.print(text);

                    // Handle images in the run
                    for (XWPFPictureData picture : pictures) {
                        if (run.getEmbeddedPictures().contains(picture)) {
                            String imageFileName = baseName + "-image-" + (imageIndex++) + "." + picture.suggestFileExtension();
                            File imageFile = new File(imageFolder, imageFileName);

                            // Write image to the image folder
                            try (FileOutputStream imageOutputStream = new FileOutputStream(imageFile)) {
                                imageOutputStream.write(picture.getData());
                            }

                            // Insert image reference in HTML
                            writer.println("<img src='" + imageFolder.getName() + "/" + imageFileName + "' alt='Image' />");
                        }
                    }
                }
                writer.println("</p>");
            }

            // End HTML content
            writer.println("</body>");
            writer.println("</html>");
        }
    }
}
