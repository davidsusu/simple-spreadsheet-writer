package hu.webarticum.simplespreadsheetwriter.example;

import hu.webarticum.simplespreadsheetwriter.Spreadsheet;

public interface Example {

    public String getLabel();
    
    public Spreadsheet create();
    
}
