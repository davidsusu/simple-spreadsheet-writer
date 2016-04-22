package hu.webarticum.simple_spreadsheet_writer.example;

import hu.webarticum.simple_spreadsheet_writer.Sheet;
import hu.webarticum.simple_spreadsheet_writer.Spreadsheet;

public class TimeTableExample implements Example {

    @Override
    public String getLabel() {
        return "Student's school timetable";
    }

    @Override
    public Spreadsheet create() {
        Spreadsheet spreadsheet = new Spreadsheet();
        
        Sheet.Format titleFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "background-color", "#777777",
            "color", "#FFFFFF",
            "font-weight", "bold",
            "font-style", "italic",
        });
        
        Sheet.Format leftFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
            "background-color", "#CCCCCC",
            "font-weight", "bold",
        });
        
        Sheet.Format basicFormat = new Sheet.Format(new String[]{
            "text-align", "center",
            "vertical-align", "middle",
        });

        Sheet sheet = spreadsheet.add("Timetable");

        sheet.getColumn(0, new Sheet.Column()).width = 10;
        sheet.getRow(0, new Sheet.Row()).height = 10;
        
        sheet.write(1, 0, "1", leftFormat);
        sheet.write(2, 0, "2", leftFormat);
        sheet.write(3, 0, "3", leftFormat);
        sheet.write(4, 0, "4", leftFormat);
        
        sheet.write(0, 1, "Monday", titleFormat);
        sheet.write(1,  1, "Math", basicFormat);
        sheet.write(2,  1, "Math", basicFormat);
        sheet.write(3,  1, "Eng", basicFormat);
        sheet.write(4,  1, "Eng", basicFormat);

        sheet.write(0, 2, "Tuesday", titleFormat);
        sheet.write(1,  2, "Music", basicFormat);
        sheet.write(2,  2, "Math", basicFormat);
        sheet.write(3,  2, "Eng", basicFormat);
        sheet.write(4,  2, "Sport", basicFormat);
        
        sheet.write(0, 3, "Wednesday", titleFormat);
        sheet.write(1,  3, "Math", basicFormat);
        sheet.write(2,  3, "Nat", basicFormat);
        sheet.write(3,  3, "Eng", basicFormat);
        sheet.write(4,  3, "Game", basicFormat);
        
        sheet.write(0, 4, "Thursday", titleFormat);
        sheet.write(1,  4, "Eng", basicFormat);
        sheet.write(2,  4, "Rel", basicFormat);
        sheet.write(3,  4, "Math", basicFormat);
        sheet.write(4,  4, "Nat", basicFormat);
        
        sheet.write(0, 5, "Friday", titleFormat);
        sheet.write(1,  5, "Eng", basicFormat);
        sheet.write(2,  5, "Eng", basicFormat);
        sheet.write(3,  5, "Sport", basicFormat);
        sheet.write(4,  5, "Game", basicFormat);
        
        return spreadsheet;
    }

}
