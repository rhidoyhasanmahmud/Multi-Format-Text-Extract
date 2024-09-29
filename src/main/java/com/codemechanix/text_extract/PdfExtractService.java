package com.codemechanix.text_extract;

import com.github.junrar.Archive;
import com.github.junrar.exception.UnsupportedRarV5Exception;
import com.github.junrar.rarfile.FileHeader;
import net.sf.sevenzipjbinding.IInArchive;
import net.sf.sevenzipjbinding.SevenZip;
import net.sf.sevenzipjbinding.impl.RandomAccessFileInStream;
import net.sf.sevenzipjbinding.simple.ISimpleInArchive;
import net.sf.sevenzipjbinding.simple.ISimpleInArchiveItem;
import org.apache.commons.compress.archivers.ArchiveEntry;
import org.apache.commons.compress.archivers.ArchiveInputStream;
import org.apache.commons.compress.archivers.ArchiveStreamFactory;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.IOUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;


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
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else if (file.getName().endsWith(".docx")) {
            try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(file.toPath()))) {
                doc.getParagraphs().forEach(paragraph -> text.append(paragraph.getText()).append("\n"));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return text.toString();
    }

    public List<String> extractFilesFromArchive(String filePath, String destinationDir) throws IOException {
        List<String> extractedFilePaths = new ArrayList<>();
        File destDir = new File(destinationDir);
        if (!destDir.exists() && !destDir.mkdirs()) {
            throw new IOException("Failed to create destination directory " + destDir);
        }
        File file = new File(filePath);
        String archiveFormat = getArchiveFormat(file.getName());

        if ("rar".equals(archiveFormat)) {
            extractRarArchive(file, destDir, extractedFilePaths);
        } else if ("7z".equals(archiveFormat)) {
            extract7zArchive(file, destDir, extractedFilePaths);
        } else {
            try (InputStream is = new BufferedInputStream(Files.newInputStream(file.toPath()));
                 ArchiveInputStream archiveInputStream = new ArchiveStreamFactory().createArchiveInputStream(archiveFormat, is)) {
                ArchiveEntry entry;
                while ((entry = archiveInputStream.getNextEntry()) != null) {
                    if (!archiveInputStream.canReadEntryData(entry)) {
                        continue;
                    }
                    File tempFile = new File(destDir, entry.getName());
                    if (entry.isDirectory()) {
                        if (!tempFile.isDirectory() && !tempFile.mkdirs()) {
                            throw new IOException("Failed to create directory " + tempFile);
                        }
                    } else {
                        File parent = tempFile.getParentFile();
                        if (!parent.isDirectory() && !parent.mkdirs()) {
                            throw new IOException("Failed to create directory " + parent);
                        }
                        try (OutputStream os = new FileOutputStream(tempFile)) {
                            IOUtils.copy(archiveInputStream, os);
                        }
                        extractedFilePaths.add(tempFile.getAbsolutePath());
                    }
                }
            } catch (Exception e) {
                throw new IOException("Error extracting files from archive: " + e.getMessage(), e);
            }
        }
        return extractedFilePaths;
    }

    private void extractRarArchive(File file, File destDir, List<String> extractedFilePaths) throws IOException {
        try (Archive archive = new Archive(file)) {
            FileHeader fileHeader;
            while ((fileHeader = archive.nextFileHeader()) != null) {
                File tempFile = new File(destDir, fileHeader.getFileName().trim());
                if (fileHeader.isDirectory()) {
                    if (!tempFile.isDirectory() && !tempFile.mkdirs()) {
                        throw new IOException("Failed to create directory " + tempFile);
                    }
                } else {
                    File parent = tempFile.getParentFile();
                    if (!parent.isDirectory() && !parent.mkdirs()) {
                        throw new IOException("Failed to create directory " + parent);
                    }
                    try (OutputStream os = new FileOutputStream(tempFile)) {
                        archive.extractFile(fileHeader, os);
                    }
                    extractedFilePaths.add(tempFile.getAbsolutePath());
                }
            }
        } catch (UnsupportedRarV5Exception e) {
            throw new IOException("RAR5 format is not supported. Please use a different tool to extract this archive.", e);
        } catch (Exception e) {
            throw new IOException("Error extracting RAR files: " + e.getMessage(), e);
        }
    }

    private void extract7zArchive(File file, File destDir, List<String> extractedFilePaths) throws IOException {
        try (RandomAccessFile randomAccessFile = new RandomAccessFile(file, "r")) {
            IInArchive inArchive = SevenZip.openInArchive(null, new RandomAccessFileInStream(randomAccessFile));
            ISimpleInArchive simpleInArchive = inArchive.getSimpleInterface();
            for (ISimpleInArchiveItem item : simpleInArchive.getArchiveItems()) {
                if (!item.isFolder()) {
                    File tempFile = new File(destDir, item.getPath());
                    File parent = tempFile.getParentFile();
                    if (!parent.isDirectory() && !parent.mkdirs()) {
                        throw new IOException("Failed to create directory " + parent);
                    }
                    try (OutputStream os = new FileOutputStream(tempFile)) {
                        final int[] hash = new int[]{0};
                        item.extractSlow(data -> {
                            try {
                                os.write(data);
                            } catch (IOException e) {
                                throw new RuntimeException(e);
                            }
                            hash[0] |= 0xFF & data[0];
                            return data.length;
                        });
                    }
                    extractedFilePaths.add(tempFile.getAbsolutePath());
                }
            }
            inArchive.close();
        } catch (Exception e) {
            throw new IOException("Error extracting 7z files: " + e.getMessage(), e);
        }
    }

    private String getArchiveFormat(String fileName) throws IOException {
        if (fileName.endsWith(".zip")) {
            return ArchiveStreamFactory.ZIP;
        } else if (fileName.endsWith(".tar")) {
            return ArchiveStreamFactory.TAR;
        } else if (fileName.endsWith(".jar")) {
            return ArchiveStreamFactory.JAR;
        } else if (fileName.endsWith(".7z")) {
            return "7z";
        } else if (fileName.endsWith(".rar")) {
            return "rar";
        } else {
            throw new IOException("Unsupported archive format: " + fileName);
        }
    }
}