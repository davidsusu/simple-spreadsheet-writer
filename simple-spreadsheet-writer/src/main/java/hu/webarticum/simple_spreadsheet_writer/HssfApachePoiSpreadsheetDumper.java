package hu.webarticum.simple_spreadsheet_writer;

import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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
        for (Map.Entry<String, String> entry: format.entrySet()) {
            String property = entry.getKey();
            String value = entry.getValue();
            if (property.equals("background-color")) {
                
            }
        }
    }
    
}
