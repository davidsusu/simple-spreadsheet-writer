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
        Sheet sheet = new Sheet();
        sheet.write(0, 0, "Sample!");
        spreadsheet.add("Timetable", sheet);
        return spreadsheet;
    }

}
