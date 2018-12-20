package hu.webarticum.simplespreadsheetwriter.example;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import hu.webarticum.simplespreadsheetwriter.Sheet;
import hu.webarticum.simplespreadsheetwriter.Spreadsheet;


public class LotOfCellsExample implements Example {

    private final int ROWS = 50;
    private final int COLUMNS = 200;
    
    private Random random = new Random(99);
    
    @Override
    public String getLabel() {
        return "Example with lot of cells";
    }

    @Override
    public Spreadsheet create() {
        Spreadsheet spreadsheet = new Spreadsheet();
        
        List<Sheet.Format> fontFormats = new ArrayList<Sheet.Format>();
        fontFormats.add(new Sheet.Format());
        fontFormats.add(new Sheet.Format("font-style", "italic"));
        fontFormats.add(new Sheet.Format("font-weight", "bold"));

        List<Sheet.Format> fontSizeFormats = new ArrayList<Sheet.Format>();
        fontSizeFormats.add(new Sheet.Format("font-size", "8pt"));
        fontSizeFormats.add(new Sheet.Format("font-size", "10pt"));
        fontSizeFormats.add(new Sheet.Format("font-size", "12pt"));
        fontSizeFormats.add(new Sheet.Format("font-size", "14pt"));
        
        List<Sheet.Format> colorFormats = new ArrayList<Sheet.Format>();
        colorFormats.add(new Sheet.Format());
        colorFormats.add(new Sheet.Format("background-color", "#FF0000", "color", "#FFFFFF"));
        colorFormats.add(new Sheet.Format("background-color", "#000000", "color", "#FFFFFF"));
        colorFormats.add(new Sheet.Format("background-color", "#CCCCCC", "color", "#000000"));
        colorFormats.add(new Sheet.Format("background-color", "#0000FF", "color", "#FFFF00"));
        colorFormats.add(new Sheet.Format("background-color", "#CC3300", "color", "#000055"));
        colorFormats.add(new Sheet.Format("background-color", "#FF00BB", "color", "#007700"));
        
        Sheet sheet = spreadsheet.add("Lot of cells");
        
        for (int row = 0; row < ROWS; row++) {
            for (int col = 0; col < COLUMNS; col++) {
                Sheet.Format cellFormat = new Sheet.Format();
                cellFormat.putAll(getRandomElement(fontFormats));
                cellFormat.putAll(getRandomElement(fontSizeFormats));
                cellFormat.putAll(getRandomElement(colorFormats));
                sheet.write(row, col, "CELL:" + (row + 1) + "," + (col + 1), cellFormat);
            }
        }
        
        return spreadsheet;
    }
    
    private <T> T getRandomElement(List<T> list) {
        if (list.isEmpty()) {
            return null;
        }
        return list.get(random.nextInt(list.size()));
    }

}
