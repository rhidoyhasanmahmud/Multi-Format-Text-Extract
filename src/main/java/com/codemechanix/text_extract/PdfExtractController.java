package com.codemechanix.text_extract;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

@RestController
public class PdfExtractController {

    @Autowired
    private PdfExtractService pdfExtractService;

     @GetMapping("/extract-text")
    public ResponseEntity<String> extractTextFromLocalPdf(@RequestParam("filePath") String encodedFilePath) {
        try {
            String filePath = URLDecoder.decode(encodedFilePath, StandardCharsets.UTF_8.name());
            String extractedText = pdfExtractService.extractTextFromFile(filePath);
            return ResponseEntity.ok(extractedText);
        } catch (IOException e) {
            return ResponseEntity.status(500).body("Error extracting text from PDF: " + e.getMessage());
        }
    }
}