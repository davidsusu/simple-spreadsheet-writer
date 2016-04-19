package hu.webarticum.simple_spreadsheet_writer.example;

import hu.webarticum.simple_spreadsheet_writer.Spreadsheet;

public interface Example {

    public String getLabel();
    
    public Spreadsheet create();
    
}
