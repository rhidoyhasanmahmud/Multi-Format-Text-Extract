package com.codemechanix.text_extract;

import java.io.File;
import java.io.IOException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.springframework.stereotype.Service;



@Service
public class PdfExtractService {

    public String extractTextFromFile(String filePath) throws IOException {
        File file = new File(filePath);
        String fileExtension = getFileExtension(file);
        System.out.println("File extension: " + fileExtension);
        switch (fileExtension) {
            case "pdf":
                return extractTextFromPdf(file);
            case "txt":
                return extractTextFromTxt(file);
            case "csv":
                return extractTextFromCsv(file);
            case "xls":
            case "xlsx":
                return extractTextFromExcel(file);
            case "doc":
            case "docx":
                return extractTextFromWord(file);
            default:
                throw new UnsupportedOperationException("File type not supported: " + fileExtension);
        }
    }

    private String getFileExtension(File file) {
        String name = file.getName();
        int lastIndexOf = name.lastIndexOf(".");
        return lastIndexOf == -1 ? "" : name.substring(lastIndexOf + 1);
    }

    private String extractTextFromPdf(File file) throws IOException {
        try (PDDocument document = PDDocument.load(file)) {
            PDFTextStripper pdfStripper = new PDFTextStripper();
            return pdfStripper.getText(document);
        }
    }

    private String extractTextFromTxt(File file) throws IOException {
        return new String(Files.readAllBytes(Paths.get(file.getPath())));
    }

    private String extractTextFromCsv(File file) throws IOException {
        StringBuilder text = new StringBuilder();
        try (CSVParser parser = CSVParser.parse(file, java.nio.charset.StandardCharsets.UTF_8, CSVFormat.DEFAULT)) {
            for (CSVRecord record : parser) {
                for (String field : record) {
                    text.append(field).append(" ");
                }
                text.append("\n");
            }
        }
        return text.toString();
    }

    private String extractTextFromExcel(File file) throws IOException {
        StringBuilder text = new StringBuilder();
        try (Workbook workbook = WorkbookFactory.create(file)) {
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        text.append(cell.toString()).append(" ");
                    }
                    text.append("\n");
                }
            }
        }
        return text.toString();
    }

    private String extractTextFromWord(File file) throws IOException {
        StringBuilder text = new StringBuilder();
        System.out.println("File name: " + file.getName());
        if (file.getName().endsWith(".doc")) {
            try (HWPFDocument doc = new HWPFDocument(Files.newInputStream(file.toPath()));
            WordExtractor extractor = new WordExtractor(doc)) {
                text.append(extractor.getText());
            }catch (Exception e) {
                e.printStackTrace();
            }
        } else if (file.getName().endsWith(".docx")) {
            try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(file.toPath()))) {
                doc.getParagraphs().forEach(paragraph -> text.append(paragraph.getText()).append("\n"));
            }catch (Exception e) {
                e.printStackTrace();
            }
        }
        return text.toString();
    }
}