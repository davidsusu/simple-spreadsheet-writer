package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

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
        // XXX
        System.out.println("START ODS DUMP...");
        
        SpreadsheetDocument outputDocument;
        try {
            outputDocument = SpreadsheetDocument.newSpreadsheetDocument();
        } catch (Exception e) {
            throw new IOException(e);
        }
        

        // XXX
        System.out.println("SAVING ODS...");
        for (Spreadsheet.Page page: spreadsheet) {
            Table outputTable = outputDocument.appendSheet(page.label);
            for (Sheet.CellEntry entry: page.sheet) {
                org.odftoolkit.simple.table.Row outputRow = outputTable.getRowByIndex(entry.rowIndex);
                org.odftoolkit.simple.table.Cell outputCell = outputRow.getCellByIndex(entry.columnIndex);
                outputCell.setDisplayText(entry.cell.text);
            }
            
            
            
            /*
            org.apache.poi.ss.usermodel.Sheet sheetDumper = workbook.createSheet(page.label);
            Integer previousRowIndex = null;
            for (Sheet.CellEntry entry: page.sheet) {
                org.apache.poi.ss.usermodel.Row rowDumper;
                if (previousRowIndex == null || entry.rowIndex != previousRowIndex) {
                    rowDumper = sheetDumper.createRow(entry.rowIndex);
                } else {
                    rowDumper = sheetDumper.getRow(entry.rowIndex);
                }
                org.apache.poi.ss.usermodel.Cell cellDumper = rowDumper.createCell(entry.columnIndex);
                cellDumper.setCellValue(entry.cell.text);
                previousRowIndex = entry.rowIndex;
            }
            */
        }
        
        
        try {
            outputDocument.save(outputStram);
        } catch (Exception e) {
            throw new IOException(e);
        }
    }
    
}
