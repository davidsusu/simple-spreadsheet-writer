package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatPart;
import org.apache.poi.ss.usermodel.CellStyle;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.odfdom.type.Color;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.style.Border;
import org.odftoolkit.simple.style.Font;
import org.odftoolkit.simple.style.StyleTypeDefinitions.CellBordersType;
import org.odftoolkit.simple.style.StyleTypeDefinitions.FontStyle;
import org.odftoolkit.simple.style.StyleTypeDefinitions.HorizontalAlignmentType;
import org.odftoolkit.simple.style.StyleTypeDefinitions.SupportedLinearMeasure;
import org.odftoolkit.simple.style.StyleTypeDefinitions.VerticalAlignmentType;
import org.odftoolkit.simple.table.Table;

import hu.webarticum.simple_spreadsheet_writer.Sheet.Row;

public class OdfToolkitSpreadsheetDumper implements SpreadsheetDumper {

    @Override
    public String getDefaultExtension() {
        return "ods";
    }
    
    @Override
    public void dump(Spreadsheet spreadheet, File file) throws IOException {
        dump(spreadheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadsheet, OutputStream outputStram) throws IOException {
        SpreadsheetDocument outputDocument;
        try {
            outputDocument = SpreadsheetDocument.newSpreadsheetDocument();
            int count = outputDocument.getSheetCount();
            for (int i = count - 1; i >= 0; i--) {
                outputDocument.removeSheet(i);
            }
        } catch (Exception e) {
            throw new IOException(e);
        }
        
        for (Spreadsheet.Page page: spreadsheet) {
            Sheet sheet = page.sheet;
            if (sheet.hasNegative()) {
                sheet = new Sheet(sheet);
                sheet.moveToNonNegative();
            }
            Table outputTable = outputDocument.appendSheet(page.label);
            for (Integer rowIndex: sheet.getRowIndexes()) {
                Row row = sheet.getRow(rowIndex);
                org.odftoolkit.simple.table.Row outputRow = outputTable.getRowByIndex(rowIndex);
                if (row.height > 0) {
                    outputRow.setHeight(row.height, false);
                }
            }
            for (Integer columnIndex: sheet.getColumnIndexes()) {
                Sheet.Column column = sheet.getColumn(columnIndex);
                org.odftoolkit.simple.table.Column outputColumn = outputTable.getColumnByIndex(columnIndex);
                if (column.width > 0) {
                    outputColumn.setWidth(column.width);
                }
            }
            Iterator<Sheet.CellEntry> iterator = sheet.iterator(Sheet.ITERATOR_COMBINED);
            while (iterator.hasNext()) {
                Sheet.CellEntry cellEntry = iterator.next();
                org.odftoolkit.simple.table.Row outputRow = outputTable.getRowByIndex(cellEntry.rowIndex);
                org.odftoolkit.simple.table.Cell outputCell = outputRow.getCellByIndex(cellEntry.columnIndex);
                outputCell.setDisplayText(cellEntry.cell.text);
                Sheet.Format computedFormat = cellEntry.getComputedFormat();
                applyFormat(outputCell, computedFormat);
            }
            for (Sheet.Range mergeRange: sheet.merges) {
                outputTable.getCellRangeByPosition(
                    mergeRange.columnIndex1, mergeRange.rowIndex1,
                    mergeRange.columnIndex2, mergeRange.rowIndex2
                ).merge();
            }
        }
        
        try {
            outputDocument.save(outputStram);
        } catch (Exception e) {
            throw new IOException(e);
        }
    }
    
    private void applyFormat(org.odftoolkit.simple.table.Cell outputCell, Sheet.Format format) {
        for (Map.Entry<String, String> entry: format.entrySet()) {
            String property = entry.getKey();
            String value = entry.getValue();
            if (property.equals("background-color")) {
                outputCell.setCellBackgroundColor(new Color(value));
            } else if (property.equals("color")) {
                Font font = getFont(outputCell);
                font.setColor(new Color(value));
                outputCell.setFont(font);
            } else if (property.equals("font-style")) {
                Font font = getFont(outputCell);
                boolean isBold = (
                    font.getFontStyle().equals(FontStyle.BOLD) ||
                    font.getFontStyle().equals(FontStyle.BOLDITALIC)
                );
                font.setFontStyle(isBold ? FontStyle.BOLDITALIC : FontStyle.ITALIC);
                outputCell.setFont(font);
            } else if (property.equals("font-weight")) {
                Font font = getFont(outputCell);
                boolean isItalic = (
                    font.getFontStyle().equals(FontStyle.ITALIC) ||
                    font.getFontStyle().equals(FontStyle.BOLDITALIC)
                );
                font.setFontStyle(isItalic ? FontStyle.BOLDITALIC : FontStyle.BOLD);
                outputCell.setFont(font);
            } else if (property.equals("text-align")) {                
                outputCell.setHorizontalAlignment(getHorizontalAligment(value));
            } else if (property.equals("vertical-align")) {
                outputCell.setVerticalAlignment(getVerticalAlignment(value));
            } else if (property.equals("border-top")) {
                outputCell.setBorders(CellBordersType.TOP, getBorder(value));
            } else if (property.equals("border-right")) {
                outputCell.setBorders(CellBordersType.RIGHT, getBorder(value));
            } else if (property.equals("border-bottom")) {
                outputCell.setBorders(CellBordersType.BOTTOM, getBorder(value));
            } else if (property.equals("border-left")) {
                outputCell.setBorders(CellBordersType.LEFT, getBorder(value));
            } // TODO
        }
    }

    private Font getFont(org.odftoolkit.simple.table.Cell outputCell) {
        Font font = outputCell.getFont();
        if (font == null) {
            font = outputCell.getStyleHandler().getFont(SpreadsheetDocument.ScriptType.WESTERN);
        }
        // XXX
        if (font.getSize() == 0) {
            font.setSize(10);
        }
        return font;
    }
    
    private TreeMap<String, HorizontalAlignmentType> horizontalAlignmentMap = null;
    private HorizontalAlignmentType getHorizontalAligment(String value) {
        if (horizontalAlignmentMap == null) {
            horizontalAlignmentMap = new TreeMap<String, HorizontalAlignmentType>();
            horizontalAlignmentMap.put("left", HorizontalAlignmentType.LEFT);
            horizontalAlignmentMap.put("center", HorizontalAlignmentType.CENTER);
            horizontalAlignmentMap.put("right", HorizontalAlignmentType.RIGHT);
            horizontalAlignmentMap.put("justify", HorizontalAlignmentType.JUSTIFY);
        }
        if (horizontalAlignmentMap.containsKey(value)) {
            return horizontalAlignmentMap.get(value);
        } else {
            return HorizontalAlignmentType.DEFAULT;
        }
    }

    private TreeMap<String, VerticalAlignmentType> verticalAlignmentMap = null;
    private VerticalAlignmentType getVerticalAlignment(String value) {
        if (verticalAlignmentMap == null) {
            verticalAlignmentMap = new TreeMap<String, VerticalAlignmentType>();
            verticalAlignmentMap.put("top", VerticalAlignmentType.TOP);
            verticalAlignmentMap.put("middle", VerticalAlignmentType.MIDDLE);
            verticalAlignmentMap.put("bottom", VerticalAlignmentType.BOTTOM);
        }
        if (verticalAlignmentMap.containsKey(value)) {
            return verticalAlignmentMap.get(value);
        } else {
            return VerticalAlignmentType.DEFAULT;
        }
    }

    private Border getBorder(String value) {
        String[] tokens = value.split(" ");
        int size = Integer.parseInt(tokens[0].replaceAll("pt$", ""));
        Color color = new Color(tokens[2]);
        SupportedLinearMeasure measure = SupportedLinearMeasure.PT;
        return new Border(color, size, measure);
    }
    
}
