package com.example.controller;

import com.example.service.WordDocumentService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@Controller
public class DocumentController {
    private static final Logger logger = LoggerFactory.getLogger(DocumentController.class);

    @Autowired
    private WordDocumentService wordDocumentService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/process")
    public ResponseEntity<ByteArrayResource> processDocument(
            @RequestParam("file") MultipartFile file,
            @RequestParam Map<String, String> replacements,
            @RequestParam(defaultValue = "0") int additionalRows) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().build();
        }

        try {
            // Remove the file parameter from replacements map
            replacements.remove("file");
            
            byte[] processedDoc = wordDocumentService.processDocument(file, replacements, additionalRows);

            ByteArrayResource resource = new ByteArrayResource(processedDoc);
            String filename = "processed_" + file.getOriginalFilename();

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                    .header(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate")
                    .header(HttpHeaders.PRAGMA, "no-cache")
                    .header(HttpHeaders.EXPIRES, "0")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                    .contentLength(processedDoc.length)
                    .body(resource);
        } catch (Exception e) {
            String errorMessage = e.getMessage() != null ? e.getMessage() : "An error occurred while processing the document";
            logger.error("Error processing document: {}", errorMessage, e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .contentType(MediaType.TEXT_PLAIN)
                    .body(new ByteArrayResource(errorMessage.getBytes()));
        }
    }
}
