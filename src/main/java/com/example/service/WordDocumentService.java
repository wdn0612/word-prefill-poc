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
import java.util.HashMap;
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
                // Check if this is the "Related Party" table
                if (isRelatedPartyTable(table)) {
                    // Only add additional rows to Related Party tables
                    processRelatedPartyTable(table, replacements, additionalRows);
                } else {
                    // Don't add additional rows to other tables
                    processTable(table, replacements, 0);
                }
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

    private boolean isRelatedPartyTable(XWPFTable table) {
        // Check if this table has text that identifies it as the "Related Party" table
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    String text = paragraph.getText();
                    if (text != null && text.contains("Related Party")) {
                        logger.debug("Found Related Party table");
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private void processRelatedPartyTable(XWPFTable mainTable, Map<String, String> replacements, int additionalRows) {
        logger.debug("Processing Related Party table with {} rows", mainTable.getRows().size());
        
        // First process the main table with replacements
        for (XWPFTableRow row : mainTable.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                replaceInCell(cell, replacements);
                
                // Look for nested tables
                List<XWPFTable> nestedTables = cell.getTables();
                if (!nestedTables.isEmpty()) {
                    // Process each nested table
                    for (XWPFTable nestedTable : nestedTables) {
                        processNestedRelatedPartyTable(nestedTable, replacements, additionalRows);
                    }
                }
            }
        }
    }
    
    private void processNestedRelatedPartyTable(XWPFTable nestedTable, Map<String, String> replacements, int additionalRows) {
        logger.debug("Processing nested table in Related Party section with {} rows", nestedTable.getRows().size());
        
        // Process existing rows with replacements
        for (XWPFTableRow row : nestedTable.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                replaceInCell(cell, replacements);
            }
        }
        
        // Add additional rows with related party information if needed
        if (additionalRows > 0 && nestedTable.getNumberOfRows() >= 2) {
            // Use the 2nd row as template (index 1)
            XWPFTableRow templateRow = nestedTable.getRows().get(1);
            
            // Create a map with the specific values for related party
            Map<String, String> relatedPartyValues = new HashMap<>();
            relatedPartyValues.put("{relatedPartyName}", "ABC");
            relatedPartyValues.put("{relatedPartyContactNumber}", "123");
            relatedPartyValues.put("{relatedPartyShareHolding}", "20%");
            
            // First, add all the new rows
            for (int i = 0; i < additionalRows; i++) {
                XWPFTableRow newRow = nestedTable.createRow();
                copyRowSettings(templateRow, newRow);
                
                // Copy content from template cells to new row cells
                for (int j = 0; j < newRow.getTableCells().size(); j++) {
                    XWPFTableCell cell = newRow.getCell(j);
                    XWPFTableCell templateCell = templateRow.getCell(j);
                    
                    // Copy paragraphs from template cell to new cell
                    for (int k = 0; k < templateCell.getParagraphs().size(); k++) {
                        XWPFParagraph templateParagraph = templateCell.getParagraphs().get(k);
                        XWPFParagraph newParagraph;
                        
                        if (k < cell.getParagraphs().size()) {
                            newParagraph = cell.getParagraphs().get(k);
                        } else {
                            newParagraph = cell.addParagraph();
                        }
                        
                        // Copy paragraph content and style
                        copyParagraphContent(templateParagraph, newParagraph);
                    }
                }
            }
            
            // Then, replace values starting from row 2 (index 1) to the end
            for (int rowIndex = 1; rowIndex < nestedTable.getNumberOfRows(); rowIndex++) {
                XWPFTableRow row = nestedTable.getRow(rowIndex);
                
                // Replace values in each cell
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        replaceInParagraph(paragraph, relatedPartyValues);
                    }
                }
            }
        }
    }

    private void processTable(XWPFTable table, Map<String, String> replacements, int additionalRows) {
        logger.debug("Processing regular table with {} rows", table.getRows().size());
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
    
    /**
     * Copies row settings from source row to target row
     * 
     * @param source The source row to copy from
     * @param target The target row to copy to
     */
    private void copyRowSettings(XWPFTableRow source, XWPFTableRow target) {
        logger.debug("Copying row settings from template row");
        
        // Copy row properties
        if (source.getCtRow().getTrPr() != null) {
            target.getCtRow().setTrPr(source.getCtRow().getTrPr());
        }

        // Copy cell properties and content
        for (int i = 0; i < source.getTableCells().size(); i++) {
            XWPFTableCell sourceCell = source.getCell(i);
            XWPFTableCell targetCell = target.getCell(i);
            
            if (targetCell == null) {
                targetCell = target.createCell();
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
    
    /**
     * Copies content and formatting from source paragraph to target paragraph
     * 
     * @param source The source paragraph to copy from
     * @param target The target paragraph to copy to
     */
    private void copyParagraphContent(XWPFParagraph source, XWPFParagraph target) {
        // Copy alignment and spacing
        if (source.getAlignment() != null) {
            target.setAlignment(source.getAlignment());
        }
        
        // Copy the text and runs
        for (XWPFRun sourceRun : source.getRuns()) {
            XWPFRun targetRun = target.createRun();
            
            // Copy text
            if (sourceRun.getText(0) != null) {
                targetRun.setText(sourceRun.getText(0));
            }
            
            // Copy formatting
            targetRun.setBold(sourceRun.isBold());
            targetRun.setItalic(sourceRun.isItalic());
            targetRun.setUnderline(sourceRun.getUnderline());
            // Use getFontSize() with caution as it's deprecated
            if (sourceRun.getFontSize() != -1) {
                targetRun.setFontSize(sourceRun.getFontSize());
            }
            if (sourceRun.getFontFamily() != null) {
                targetRun.setFontFamily(sourceRun.getFontFamily());
            }
            if (sourceRun.getColor() != null) {
                targetRun.setColor(sourceRun.getColor());
            }
        }
    }
}
