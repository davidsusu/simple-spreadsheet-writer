package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;


public class HtmlSpreadsheetDumper implements SpreadsheetDumper {

    @Override
    public String getDefaultExtension() {
        return "html";
    }
    
    @Override
    public void dump(Spreadsheet spreadsheet, File file) throws IOException {
        dump(spreadsheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadsheet, OutputStream outputStream) throws IOException {
        // TODO
    }

}
