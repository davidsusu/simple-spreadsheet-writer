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
    public void dump(Spreadsheet spreadheet, File file) throws IOException {
        dump(spreadheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadheet, OutputStream outputStram) throws IOException {
        // TODO
    }

}
