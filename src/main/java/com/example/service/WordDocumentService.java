package com.example.service;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class WordDocumentService {
    private static final Logger logger = LoggerFactory.getLogger(WordDocumentService.class);

    public byte[] processDocument(MultipartFile template, Map<String, String> replacements, int additionalRows) throws IOException {
        logger.info("Processing document with replacements: {}", replacements);
        
        try (InputStream is = template.getInputStream();
             XWPFDocument doc = new XWPFDocument(is);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            
            // Process tables
            for (XWPFTable table : doc.getTables()) {
                processTable(table, replacements, additionalRows);
            }

            // Process paragraphs outside tables
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                replaceInParagraph(paragraph, replacements);
            }

            // Write the modified document to a byte array
            doc.write(outputStream);
            outputStream.flush();
            return outputStream.toByteArray();
        }
    }

    private void replaceInCell(XWPFTableCell cell, Map<String, String> replacements) {
        logger.debug("Processing cell content");
        
        // Process all paragraphs in the cell
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            replaceInParagraph(paragraph, replacements);
        }
        
        // Process any nested tables
        for (XWPFTable nestedTable : cell.getTables()) {
            processTable(nestedTable, replacements, 0); // Don't add rows to nested tables
        }
    }

    private void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> replacements) {
        try {
            String paragraphText = paragraph.getText();
            logger.debug("Original paragraph text: {}", paragraphText);

            // Check if any replacements are needed
            boolean needsReplacement = false;
            for (String key : replacements.keySet()) {
                if (paragraphText.contains(key)) {
                    needsReplacement = true;
                    break;
                }
            }

            if (!needsReplacement) {
                return;
            }

            // Store the original style
            CTPPr originalStyle = null;
            if (paragraph.getCTP().getPPr() != null) {
                originalStyle = paragraph.getCTP().getPPr();
            }

            // Copy and modify text with replacements
            String newText = paragraphText;
            for (Map.Entry<String, String> entry : replacements.entrySet()) {
                newText = newText.replace(entry.getKey(), entry.getValue());
                logger.debug("Replacing {} with {}", entry.getKey(), entry.getValue());
            }

            // Clear existing runs
            while (paragraph.getRuns().size() > 0) {
                paragraph.removeRun(0);
            }

            // Create a new run with the replaced text
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(newText);

            // Copy formatting from the original style if available
            if (originalStyle != null) {
                paragraph.getCTP().setPPr(originalStyle);
            }

            logger.debug("Successfully replaced text in paragraph");
        } catch (Exception e) {
            logger.error("Error replacing text in paragraph", e);
        }
    }

    private void processTable(XWPFTable table, Map<String, String> replacements, int additionalRows) {
        logger.debug("Processing table with {} rows", table.getRows().size());
        // Process existing rows
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                replaceInCell(cell, replacements);
            }
        }

        // Add additional rows if needed
        if (additionalRows > 0) {
            XWPFTableRow templateRow = table.getRows().get(table.getNumberOfRows() - 1);
            for (int i = 0; i < additionalRows; i++) {
                XWPFTableRow newRow = table.createRow();
                copyRowSettings(templateRow, newRow);
            }
        }
    }

    private void copyRowSettings(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
        logger.debug("Copying row settings from template row");
        
        // Copy row properties
        if (sourceRow.getCtRow().getTrPr() != null) {
            targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        }

        // Copy cell properties and content
        for (int i = 0; i < sourceRow.getTableCells().size(); i++) {
            XWPFTableCell sourceCell = sourceRow.getCell(i);
            XWPFTableCell targetCell = targetRow.getCell(i);
            
            if (targetCell == null) {
                targetCell = targetRow.createCell();
            }

            // Copy cell properties
            if (sourceCell.getCTTc().getTcPr() != null) {
                targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            }

            // Copy cell content and formatting
            for (XWPFParagraph sourcePara : sourceCell.getParagraphs()) {
                XWPFParagraph targetPara = targetCell.addParagraph();
                // Copy paragraph properties
                if (sourcePara.getCTP().getPPr() != null) {
                    targetPara.getCTP().setPPr(sourcePara.getCTP().getPPr());
                }
            }
        }
    }
}
