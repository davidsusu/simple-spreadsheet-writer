package hu.webarticum.simplespreadsheetwriter;

import java.io.Serializable;
import java.util.LinkedList;

public class Spreadsheet extends LinkedList<Spreadsheet.Page> {
    
    private static final long serialVersionUID = -5236028318494132091L;

    public void add(String label, Sheet sheet) {
        this.add(new Page(label, sheet));
    }

    public Sheet add(String label) {
        Sheet sheet = new Sheet();
        this.add(label, sheet);
        return sheet;
    }
    
    public class Page implements Serializable {
        
        private static final long serialVersionUID = 7505335497467115377L;

        public String label;
        
        public Sheet sheet;
        
        public Page(String label, Sheet sheet) {
            this.label = label;
            this.sheet = sheet;
        }
        
    }
    
}
