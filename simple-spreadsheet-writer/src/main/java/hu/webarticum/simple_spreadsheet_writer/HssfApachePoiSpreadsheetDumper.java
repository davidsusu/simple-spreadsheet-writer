package hu.webarticum.simple_spreadsheet_writer;

import org.apache.poi.ss.usermodel.Workbook;
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
    
}
