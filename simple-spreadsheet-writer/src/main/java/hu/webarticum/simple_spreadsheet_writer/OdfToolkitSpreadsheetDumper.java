package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatPart;
import org.apache.poi.ss.usermodel.CellStyle;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.odfdom.type.Color;
import org.odftoolkit.simple.SpreadsheetDocument;
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
                // TODO
            } else if (property.equals("font-weight")) {
                // TODO
            } // TODO
        }
    }
    
}
