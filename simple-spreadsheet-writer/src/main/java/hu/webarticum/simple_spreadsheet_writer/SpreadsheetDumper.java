package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

public interface SpreadsheetDumper {

    public String getDefaultExtension();
    
    public void dump(Spreadsheet spreadheet, File file) throws IOException;
    
    public void dump(Spreadsheet spreadheet, OutputStream outputStram) throws IOException;
    
}
