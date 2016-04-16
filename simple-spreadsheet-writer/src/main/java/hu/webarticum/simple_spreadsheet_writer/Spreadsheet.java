package hu.webarticum.simple_spreadsheet_writer;

import java.util.LinkedList;

public class Spreadsheet extends LinkedList<Spreadsheet.Page> {
    
    private static final long serialVersionUID = 1L;

    public void add(String label, Sheet sheet) {
        this.add(new Page(label, sheet));
    }

    public Sheet add(String label) {
        Sheet sheet = new Sheet();
        this.add(label, sheet);
        return sheet;
    }
    
    public class Page {
        
        public String label;
        
        public Sheet sheet;
        
        public Page(String label, Sheet sheet) {
            this.label = label;
            this.sheet = sheet;
        }
        
    }
    
}
