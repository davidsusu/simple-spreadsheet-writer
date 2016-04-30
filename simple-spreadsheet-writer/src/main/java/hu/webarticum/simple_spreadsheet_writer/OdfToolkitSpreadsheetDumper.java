package hu.webarticum.simple_spreadsheet_writer;

import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import org.odftoolkit.odfdom.type.Color;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.style.Border;
import org.odftoolkit.simple.style.Font;
import org.odftoolkit.simple.style.StyleTypeDefinitions.CellBordersType;
import org.odftoolkit.simple.style.StyleTypeDefinitions.FontStyle;
import org.odftoolkit.simple.style.StyleTypeDefinitions.HorizontalAlignmentType;
import org.odftoolkit.simple.style.StyleTypeDefinitions.LineStyle;
import org.odftoolkit.simple.style.StyleTypeDefinitions.SupportedLinearMeasure;
import org.odftoolkit.simple.style.StyleTypeDefinitions.VerticalAlignmentType;
import org.odftoolkit.simple.table.Table;

import hu.webarticum.simple_spreadsheet_writer.Sheet.Row;
import hu.webarticum.simple_spreadsheet_writer.util.ColorUtil;

public class OdfToolkitSpreadsheetDumper implements SpreadsheetDumper {

    static protected final int DEFAULT_FONTSIZE = 10;

    static protected final float MM_PTS = 2.834645669291f;
    
    @Override
    public String getDefaultExtension() {
        return "ods";
    }
    
    @Override
    public void dump(Spreadsheet spreadsheet, File file) throws IOException {
        dump(spreadsheet, new FileOutputStream(file));
    }

    public void dump(Spreadsheet spreadsheet, OutputStream outputStream) throws IOException {
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
            for (Sheet.Range mergeRange: sheet.merges) {
                outputTable.getCellRangeByPosition(
                    mergeRange.columnIndex1, mergeRange.rowIndex1,
                    mergeRange.columnIndex2, mergeRange.rowIndex2
                ).merge();
            }
            Iterator<Sheet.CellEntry> iterator = sheet.iterator(Sheet.ITERATOR_COMBINED);
            while (iterator.hasNext()) {
                Sheet.CellEntry cellEntry = iterator.next();
                org.odftoolkit.simple.table.Row outputRow = outputTable.getRowByIndex(cellEntry.rowIndex);
                org.odftoolkit.simple.table.Cell outputCell = outputRow.getCellByIndex(cellEntry.columnIndex);
                fixCell(outputCell);
                outputCell.setStringValue(cellEntry.cell.text);
                Sheet.Format computedFormat = cellEntry.getComputedFormat();
                applyFormat(outputCell, computedFormat, cellEntry.rowIndex, cellEntry.columnIndex);
            }
            for (Integer rowIndex: sheet.getRowIndexes()) {
                Row row = sheet.getRow(rowIndex);
                org.odftoolkit.simple.table.Row outputRow = outputTable.getRowByIndex(rowIndex);
                if (row.height > 0) {
                    outputRow.setHeight(row.height, false);
                } else if (row.height == (-1)) {
                    outputRow.setUseOptimalHeight(true);
                }
            }
            for (Integer columnIndex: sheet.getColumnIndexes()) {
                Sheet.Column column = sheet.getColumn(columnIndex);
                org.odftoolkit.simple.table.Column outputColumn = outputTable.getColumnByIndex(columnIndex);
                if (column.width > 0) {
                    outputColumn.setWidth(column.width);
                } else if (column.width == (-1)) {
                    outputColumn.setUseOptimalWidth(true);
                    
                    // XXX setUseOptimalWidth does not work
                    double maxWidth = outputColumn.getWidth();
                    for (int i = outputColumn.getCellCount() -1; i >= 0; i--) {
                        org.odftoolkit.simple.table.Cell outputCell = outputColumn.getCellByIndex(i);
                        String value = outputCell.getStringValue();
                        String[] lines;
                        if (value.contains("\n")) {
                            lines = value.split("\\n");
                        } else {
                            lines = new String[]{value};
                        }
                        for (String line: lines) {
                            double lineWidth = createAwtFont(outputCell.getFont()).getStringBounds(
                                line, new FontRenderContext(new AffineTransform(), true, true)
                            ).getWidth() / MM_PTS;
                            if (lineWidth > maxWidth) {
                                maxWidth = lineWidth;
                            }
                        }
                    }
                    outputColumn.setWidth(maxWidth);
                }
            }
        }
        
        try {
            outputDocument.save(outputStream);
        } catch (Exception e) {
            throw new IOException(e);
        } finally {
            outputStream.close();
        }
    }
    
    // XXX cells created potentially bad
    protected void fixCell(org.odftoolkit.simple.table.Cell outputCell) {
        outputCell.setCellBackgroundColor((Color)null);
        outputCell.setHorizontalAlignment(HorizontalAlignmentType.DEFAULT);
        outputCell.setVerticalAlignment(VerticalAlignmentType.DEFAULT);
        outputCell.setTextWrapped(false);
        outputCell.setBorders(CellBordersType.TOP, null);
        outputCell.setBorders(CellBordersType.RIGHT, null);
        outputCell.setBorders(CellBordersType.BOTTOM, null);
        outputCell.setBorders(CellBordersType.LEFT, null);
        {
            Font font = getFont(outputCell);
            font.setColor(getColor("#000000"));
            font.setFontStyle(FontStyle.REGULAR);
            font.setSize(DEFAULT_FONTSIZE);
            outputCell.setFont(font);
        }
        outputCell.removeContent();
    }
    
    protected void applyFormat(org.odftoolkit.simple.table.Cell outputCell, Sheet.Format format, int rowIndex, int columnIndex) {
        for (Map.Entry<String, String> entry: format.entrySet()) {
            String property = entry.getKey();
            String value = entry.getValue();
            if (property.equals("background-color")) {
                outputCell.setCellBackgroundColor(getColor(value));
            } else if (property.equals("color")) {
                Font font = getFont(outputCell);
                font.setColor(getColor(value));
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
            } else if (property.equals("white-space")) {
                outputCell.setTextWrapped(value.matches("pre\\b.*"));
            } else if (property.equals("border-top")) {
                outputCell.setBorders(CellBordersType.TOP, getBorder(value));
            } else if (property.equals("border-right")) {
                outputCell.setBorders(CellBordersType.RIGHT, getBorder(value));
            } else if (property.equals("border-bottom")) {
                outputCell.setBorders(CellBordersType.BOTTOM, getBorder(value));
            } else if (property.equals("border-left")) {
                outputCell.setBorders(CellBordersType.LEFT, getBorder(value));
            } else if (property.equals("font-size")) {
                Font font = getFont(outputCell);
                double size;
                if (value.endsWith("pt")) {
                    size = Double.parseDouble(value.replaceAll("pt$", ""));
                } else if (value.endsWith("%")) {
                    size = DEFAULT_FONTSIZE * Double.parseDouble(value.replaceAll("%$", "")) / 100;
                } else {
                    size = DEFAULT_FONTSIZE;
                }
                font.setSize(size);
                outputCell.setFont(font);
            } // TODO
        }
    }

    protected Font getFont(org.odftoolkit.simple.table.Cell outputCell) {
        Font font = outputCell.getFont();
        if (font == null) {
            font = outputCell.getStyleHandler().getFont(SpreadsheetDocument.ScriptType.WESTERN);
        }
        // XXX
        if (font.getSize() == 0) {
            font.setSize(DEFAULT_FONTSIZE);
        }
        return font;
    }
    
    protected TreeMap<String, HorizontalAlignmentType> horizontalAlignmentMap = null;
    protected HorizontalAlignmentType getHorizontalAligment(String value) {
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

    protected TreeMap<String, VerticalAlignmentType> verticalAlignmentMap = null;
    protected VerticalAlignmentType getVerticalAlignment(String value) {
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

    protected Border getBorder(String value) {
        String[] tokens = value.split(" ");
        double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
        Color color = getColor(tokens[2]);
        SupportedLinearMeasure measure = SupportedLinearMeasure.PT;
        
        @SuppressWarnings("unused")
        LineStyle lineStyle = LineStyle.SOLID;
        if (tokens[1].equals("dotted")) {
            lineStyle = LineStyle.DOTTED;
        } else if (tokens[1].equals("dashed")) {
            lineStyle = LineStyle.DASH;
        }
        
        Border border = new Border(color, size, measure);
        
        // lineStyle???
        
        return border;
    }
    
    protected Color getColor(String value) {
        int[] parts = ColorUtil.parseColor(value);
        return new Color(parts[0], parts[1], parts[2]);
    }
    
    protected java.awt.Font createAwtFont(Font font) {
        String fontName = "Arial";
        FontStyle inputFontStyle = font.getFontStyle();
        int fontStyle;
        switch (inputFontStyle) {
            case BOLD:
                fontStyle = java.awt.Font.BOLD;
                break;
            case ITALIC:
                fontStyle = java.awt.Font.ITALIC;
                break;
            case BOLDITALIC:
                fontStyle = (java.awt.Font.BOLD | java.awt.Font.ITALIC);
                break;
            default:
                fontStyle = java.awt.Font.PLAIN;
        }
        int fontSize = (int)font.getSize();
        return new java.awt.Font(fontName, fontStyle, fontSize);
    }
    
}
