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

            // Extract the document text
            WordExtractor extractor = new WordExtractor(doc);
            String[] paragraphs = extractor.getParagraphText();

            // Begin writing HTML content
            writer.println("<html>");
            writer.println("<head><style>body { font-family: Arial, sans-serif; }</style></head>");
            writer.println("<body>");

            // Extract pictures from the document
            PicturesTable picturesTable = doc.getPicturesTable();
            List<Picture> pictures = picturesTable.getAllPictures();
            int imageIndex = 0;

            // Loop through paragraphs and convert them to HTML paragraphs
            for (String paragraph : paragraphs) {
                writer.println("<p>" + paragraph.trim() + "</p>");
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
