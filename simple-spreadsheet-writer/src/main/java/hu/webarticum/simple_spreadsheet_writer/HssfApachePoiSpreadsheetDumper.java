package hu.webarticum.simple_spreadsheet_writer;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import hu.webarticum.simple_spreadsheet_writer.util.ColorUtil;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HssfApachePoiSpreadsheetDumper extends ApachePoiSpreadsheetDumper {

    @Override
    public String getDefaultExtension() {
        return "xls";
    }
    
    @Override
    protected Workbook createWorkbook() {
        return new HSSFWorkbook();
    }

    @Override
    protected void applyProblematicFormat(Workbook outputWorkbook, org.apache.poi.ss.usermodel.Cell outputCell, Sheet.Format format) {
        HSSFWorkbook workbook = (HSSFWorkbook)outputWorkbook;
        HSSFCell cell = (HSSFCell)outputCell;
        HSSFCellStyle cellStyle = cell.getCellStyle();
        HSSFFont font = null;
        for (Map.Entry<String, String> entry: format.entrySet()) {
            String property = entry.getKey();
            String value = entry.getValue();
            if (property.equals("background-color")) {
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(getColor(workbook, value));
            } else if (property.equals("color")) {
                if (font == null) {
                    font = workbook.createFont();
                }
                font.setColor(getColor(workbook, value));
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
                short size = font.getFontHeight();
                if (value.endsWith("pt")) {
                    size = (short)(Double.parseDouble(value.replaceAll("pt$", "")) * 20);
                } else if (value.endsWith("%")) {
                    size = (short)(size * Short.parseShort(value.replaceAll("%$", "")) / 100);
                }
                font.setFontHeight(size);
            } else if (property.equals("border-top")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderTop((size > 0.9) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setTopBorderColor(getColor(workbook, tokens[2]));
            } else if (property.equals("border-right")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderRight((size > 0.9) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setRightBorderColor(getColor(workbook, tokens[2]));
            } else if (property.equals("border-bottom")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderBottom((size > 0.9) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setBottomBorderColor(getColor(workbook, tokens[2]));
            } else if (property.equals("border-left")) {
                String[] tokens = value.split(" ");
                // XXX
                double size = Double.parseDouble(tokens[0].replaceAll("pt$", ""));
                cellStyle.setBorderLeft((size > 0.9) ? CellStyle.BORDER_MEDIUM : CellStyle.BORDER_THIN);
                cellStyle.setLeftBorderColor(getColor(workbook, tokens[2]));
            }
        }
        if (font != null) {
            cellStyle.setFont(font);
        }
    }
    
    private Map<String, Short> colorIndexMap = null;
    private short colorIndex = 0x8;
    private short getColor(HSSFWorkbook workbook, String value) {
        if (colorIndexMap == null) {
            colorIndexMap = new HashMap<String, Short>();
        } else if (colorIndexMap.containsKey(value)) {
            return colorIndexMap.get(value);
        }
        short index = (colorIndex++);
        int[] parts = ColorUtil.parseColor(value);
        HSSFPalette palette = workbook.getCustomPalette();
        // XXX
        palette.setColorAtIndex(
            index,
            (byte)parts[0],
            (byte)parts[1],
            (byte)parts[2]
        );
        colorIndexMap.put(value, index);
        return index;
    }
    
}
