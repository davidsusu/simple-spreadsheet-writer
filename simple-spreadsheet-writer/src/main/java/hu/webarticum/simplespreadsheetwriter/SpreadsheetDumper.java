package hu.webarticum.simplespreadsheetwriter;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

public interface SpreadsheetDumper {

    public String getDefaultExtension();
    
    public void dump(Spreadsheet spreadsheet, File file) throws IOException;
    
    public void dump(Spreadsheet spreadsheet, OutputStream outputStream) throws IOException;
    
}
