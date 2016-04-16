package hu.webarticum.simple_spreadsheet_writer;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XssfApachePoiSpreadsheetDumper extends ApachePoiSpreadsheetDumper {

    @Override
    protected Workbook createWorkbook() {
        return new XSSFWorkbook();
    }
    
}