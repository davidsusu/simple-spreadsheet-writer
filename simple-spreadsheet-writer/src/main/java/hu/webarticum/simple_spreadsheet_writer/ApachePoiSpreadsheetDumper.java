package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;

abstract public class ApachePoiSpreadsheetDumper implements SpreadsheetDumper {

    @Override
    public void dump(Spreadsheet spreadheet, File file) throws IOException {
        dump(spreadheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadsheet, OutputStream outputStram) throws IOException {
        Workbook outputWorkbook = createWorkbook();
        
        // XXX
        for (Spreadsheet.Page page: spreadsheet) {
            org.apache.poi.ss.usermodel.Sheet outputSheet = outputWorkbook.createSheet(page.label);
            Integer previousRowIndex = null;
            for (Sheet.CellEntry entry: page.sheet) {
                org.apache.poi.ss.usermodel.Row outputRow;
                if (previousRowIndex == null || entry.rowIndex != previousRowIndex) {
                    outputRow = outputSheet.createRow(entry.rowIndex);
                } else {
                    outputRow = outputSheet.getRow(entry.rowIndex);
                }
                org.apache.poi.ss.usermodel.Cell outputCell = outputRow.createCell(entry.columnIndex);
                outputCell.setCellValue(entry.cell.text);
                previousRowIndex = entry.rowIndex;
            }
        }
        
        outputWorkbook.write(outputStram);
    }
    
    abstract protected Workbook createWorkbook();
    
}
