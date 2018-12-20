package hu.webarticum.simplespreadsheetwriter.example;

import hu.webarticum.simplespreadsheetwriter.Sheet;
import hu.webarticum.simplespreadsheetwriter.Spreadsheet;

public class TimeTableExample implements Example {

    @Override
    public String getLabel() {
        return "Student's school timetable";
    }

    @Override
    public Spreadsheet create() {
        Spreadsheet spreadsheet = new Spreadsheet();

        Sheet.Format mainTitleFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "background-color", "#FFFC99",
            "color", "#553300",
            "font-weight", "bold",
            "border-top", "1pt solid #333333",
            "border-right", "1pt solid #333333",
            "border-bottom", "1pt solid #333333",
            "border-left", "1pt solid #333333",
            "font-size", "15pt",
        });

        Sheet.Format cornerFormat = new Sheet.Format(new String[]{
            "border-top", "0.5pt solid #333333",
            "border-right", "0.5pt solid #333333",
            "border-bottom", "0.5pt solid #333333",
            "border-left", "0.5pt solid #333333",
        });

        Sheet.Format titleFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "background-color", "#777777",
            "color", "#FFFFFF",
            "font-weight", "bold",
            "font-style", "italic",
            "font-size", "110%",
            "border-top", "0.5pt solid #333333",
            "border-right", "0.5pt solid #333333",
            "border-bottom", "0.5pt solid #333333",
            "border-left", "0.5pt solid #333333",
        });
        
        Sheet.Format leftFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "background-color", "#CCCCCC",
            "font-weight", "bold",
            "font-size", "110%",
            "border-top", "0.5pt solid #333333",
            "border-right", "0.5pt solid #333333",
            "border-bottom", "0.5pt solid #333333",
            "border-left", "0.5pt solid #333333",
        });
        
        Sheet.Format basicFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "border-top", "0.5pt solid #333333",
            "border-right", "0.5pt solid #333333",
            "border-bottom", "0.5pt solid #333333",
            "border-left", "0.5pt solid #333333",
        });

        Sheet sheet = spreadsheet.add("Timetable");

        sheet.getColumn(0, new Sheet.Column()).width = 10;
        sheet.getRow(0, new Sheet.Row()).height = 20;
        sheet.getRow(1, new Sheet.Row()).height = 10;
        
        sheet.write(0, 0, "Sample timetable", mainTitleFormat);
        sheet.merges.add(new Sheet.Range(0, 0, 0, 5));
        
        sheet.write(1,  0, "", cornerFormat);
        
        sheet.write(2, 0, "1", leftFormat);
        sheet.write(3, 0, "2", leftFormat);
        sheet.write(4, 0, "3", leftFormat);
        sheet.write(5, 0, "4", leftFormat);
        
        sheet.write(1, 1, "Monday", titleFormat);
        sheet.write(2,  1, "Math", basicFormat);
        sheet.write(3,  1, "Math", basicFormat);
        sheet.write(4,  1, "Eng", basicFormat);
        sheet.write(5,  1, "Eng", basicFormat);

        sheet.write(1, 2, "Tuesday", titleFormat);
        sheet.write(2,  2, "Music", basicFormat);
        sheet.write(3,  2, "Math", basicFormat);
        sheet.write(4,  2, "Eng", basicFormat);
        sheet.write(5,  2, "Sport", basicFormat);
        
        sheet.write(1, 3, "Wednesday", titleFormat);
        sheet.write(2,  3, "Math", basicFormat);
        sheet.write(3,  3, "Nat", basicFormat);
        sheet.write(4,  3, "Eng", basicFormat);
        sheet.write(5,  3, "Game", basicFormat);
        
        sheet.write(1, 4, "Thursday", titleFormat);
        sheet.write(2,  4, "Eng", basicFormat);
        sheet.write(3,  4, "Rel", basicFormat);
        sheet.write(4,  4, "Math", basicFormat);
        sheet.write(5,  4, "Nat", basicFormat);
        
        sheet.write(1, 5, "Friday", titleFormat);
        sheet.write(2,  5, "Eng", basicFormat);
        sheet.write(3,  5, "Eng", basicFormat);
        sheet.write(4,  5, "Sport", basicFormat);
        sheet.write(5,  5, "Game", basicFormat);
        
        return spreadsheet;
    }

}
