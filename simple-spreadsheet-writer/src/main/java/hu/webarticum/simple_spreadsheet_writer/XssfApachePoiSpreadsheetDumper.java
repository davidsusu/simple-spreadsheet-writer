package hu.webarticum.simple_spreadsheet_writer;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import hu.webarticum.simple_spreadsheet_writer.util.ColorUtil;

public class XssfApachePoiSpreadsheetDumper extends ApachePoiSpreadsheetDumper {

    // XXX
    static protected final int DEFAULT_FONTSIZE = 9;
    
    @Override
    public String getDefaultExtension() {
        return "xlsx";
    }
    
    @Override
    protected Workbook createWorkbook() {
        return new XSSFWorkbook();
    }

    @Override
    protected void applyProblematicFormat(Workbook outputWorkbook, org.apache.poi.ss.usermodel.Cell outputCell, Sheet.Format format) {
        XSSFWorkbook workbook = (XSSFWorkbook)outputWorkbook;
        XSSFCell cell = (XSSFCell)outputCell;
        XSSFCellStyle cellStyle = cell.getCellStyle();
        XSSFFont font = null;
        for (Map.Entry<String, String> entry: format.entrySet()) {
            String property = entry.getKey();
            String value = entry.getValue();
            if (property.equals("background-color")) {
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(getColor(value));
            } else if (property.equals("color")) {
                if (font == null) {
                    font = workbook.createFont();
                }
                font.setColor(getColor(value));
            } else if (property.equals("font-style")) {
                if (font == null) {
                    font = workbook.createFont();
                }
                font.setItalic(value.equals("italic"));
            } else if (property.equals("font-weight")) {
                if (font == null) {
                    font = workbook.createFont();
                }
                font.setBold(value.equals("bold"));
            } else if (property.equals("font-size")) {
                if (font == null) {
                    font = workbook.createFont();
                }
                double size;
                if (value.endsWith("pt")) {
                    size = Double.parseDouble(value.replaceAll("pt$", ""));
                } else if (value.endsWith("%")) {
                    size = DEFAULT_FONTSIZE * Double.parseDouble(value.replaceAll("%$", "")) / 100;
                } else {
                    size = DEFAULT_FONTSIZE;
                }
                font.setFontHeight(size);
            } else if (property.equals("border-top")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderTop((size > 1) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setTopBorderColor(getColor(tokens[2]));
            } else if (property.equals("border-right")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderRight((size > 1) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setRightBorderColor(getColor(tokens[2]));
            } else if (property.equals("border-bottom")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderBottom((size > 1) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setBottomBorderColor(getColor(tokens[2]));
            } else if (property.equals("border-left")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderLeft((size > 1) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setLeftBorderColor(getColor(tokens[2]));
            }
            
        }
        if (font != null) {
            cellStyle.setFont(font);
        }
        cell.setCellStyle(cellStyle);
    }
    
    private XSSFColor getColor(String value) {
        int[] parts = ColorUtil.parseColor(value);
        java.awt.Color awtColor = new java.awt.Color(parts[0], parts[1], parts[2]);
        return new XSSFColor(awtColor);
    }
    
}
