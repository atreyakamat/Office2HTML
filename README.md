
# 🌟 File to HTML Converter 🌟

A versatile and efficient tool to convert `.doc`, `.docx`, `.xls`, and `.xlsx` files into clean and structured HTML documents. 🚀

---

## 📖 Overview

The **File to HTML Converter** is designed to handle various file formats, extracting content such as text, tables, and images, and presenting them in an HTML format. It supports interactive features through embedded JavaScript for enhanced functionality. With its centralized JavaScript management and clean output structure, this project is ideal for generating web-ready content from office documents.

---

## 🔑 Features

- **Multi-Format Conversion**  
  - Convert `.doc` and `.docx` files into HTML.  
  - Transform `.xls` and `.xlsx` spreadsheets into HTML.  

- **Content Preservation**  
  - Extracts and accurately displays:  
    - Text  
    - Tables  
    - Images  

- **Interactive HTML Output**  
  - Includes embedded JavaScript functionalities for seamless integration.  

- **Centralized JavaScript Management**  
  - JavaScript code is managed in a single Java class, making updates effortless.

- **Organized Output**  
  - HTML files and images are saved in a well-structured format.

---

## 📂 Project Structure

```
src/
├── DocConverter.java             # Handles DOC to HTML conversion
├── DocxConverter.java            # Handles DOCX to HTML conversion
├── ExcelConverter.java           # Handles XLS/XLSX to HTML conversion
├── JavaScriptImplementation.java # Centralized JavaScript code as strings
├── MainRun.java                  # Main entry point of the application
└── resources/                    # Contains input and output directories
```

---

## 🚀 Getting Started

### Prerequisites

- **Java Development Kit (JDK)** 8 or above.  
- A Java IDE or terminal to compile and run the project.  

### Installation

1. Clone the repository:  
   ```bash
   git clone <repository_url>
   cd <repository_folder>
   ```

2. Compile the Java files:  
   ```bash
   javac -d bin src/*.java
   ```

3. Run the program:  
   ```bash
   java -cp bin MainRun
   ```

---

## 📂 Output Details

- Converted `.html` files are saved in the specified output folder.  
- Extracted images are stored in an `images/` subfolder.  

---

## 👨‍💻 Usage

1. Run the program and provide the input file path.  
2. Specify the output directory.  
3. Enjoy the fully converted HTML file with content, images, and tables!

---

## 🧑‍💻 Contributors

Developed with ❤️ by **Atreya Kamat**.

---

