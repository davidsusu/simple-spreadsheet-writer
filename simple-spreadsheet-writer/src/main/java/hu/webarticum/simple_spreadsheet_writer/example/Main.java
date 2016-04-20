package hu.webarticum.simple_spreadsheet_writer.example;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import hu.webarticum.simple_spreadsheet_writer.HssfApachePoiSpreadsheetDumper;
import hu.webarticum.simple_spreadsheet_writer.HtmlSpreadsheetDumper;
import hu.webarticum.simple_spreadsheet_writer.OdfToolkitSpreadsheetDumper;
import hu.webarticum.simple_spreadsheet_writer.Spreadsheet;
import hu.webarticum.simple_spreadsheet_writer.SpreadsheetDumper;
import hu.webarticum.simple_spreadsheet_writer.XssfApachePoiSpreadsheetDumper;

public class Main {

    public static void main(String[] args) {
        List<Example> examples = new ArrayList<Example>();
        examples.add(new TimeTableExample());
        
        System.out.println("Examples:");
        int count = examples.size();
        for (int i = 0; i < count; i++) {
            System.out.println(" " + (i + 1) + ") " + examples.get(i).getLabel());
        }
        
        int defaultNumber = 1;
        
        int number;
        while (true) {
            System.out.print("Choose an example (1): ");
            String line;
            try {
                line = (new BufferedReader(new InputStreamReader(System.in))).readLine();
            } catch (IOException e1) {
                System.out.println("Input error!");
                continue;
            }
            if (line.isEmpty()) {
                number = defaultNumber;
                break;
            }
            try {
                number = Integer.parseInt(line);
            } catch (NumberFormatException e) {
                System.out.println("This is not an integer!");
                continue;
            }
            if (number < 1 || number > count) {
                System.out.println("Please type a number between 1 and " + count + "!");
                continue;
            }
            break;
        }
        
        Example example = examples.get(number - 1);
        System.out.println(example.getLabel() + " choosen");
        
        List<SpreadsheetDumper> dumpers = new ArrayList<SpreadsheetDumper>();
        dumpers.add(new OdfToolkitSpreadsheetDumper());
        dumpers.add(new HssfApachePoiSpreadsheetDumper());
        dumpers.add(new XssfApachePoiSpreadsheetDumper());
        dumpers.add(new HtmlSpreadsheetDumper());
        
        Map<String, SpreadsheetDumper> dumperMap = new LinkedHashMap<String, SpreadsheetDumper>();
        for (SpreadsheetDumper dumper: dumpers) {
            dumperMap.put(dumper.getDefaultExtension(), dumper);
        }
        
        String defaultType = "ods";

        Spreadsheet spreadsheet = null;
        while (true) {
            System.out.print("Output file: ");
            String line;
            try {
                line = (new BufferedReader(new InputStreamReader(System.in))).readLine();
            } catch (IOException e1) {
                System.out.println("Input error!");
                continue;
            }
            String type = defaultType;
            int perPos = line.lastIndexOf('/');
            int dotPos = line.lastIndexOf('.');
            if (dotPos > perPos) {
                type = line.substring(dotPos + 1);
            } else {
                System.out.println("No extension detected, ." + defaultType + " format will be used.");
            }
            if (!dumperMap.containsKey(type)) {
                System.out.println("Unknown extension!");
                continue;
            }
            
            SpreadsheetDumper dumper = dumperMap.get(type);
            
            if (spreadsheet == null) {
                spreadsheet = example.create();
            }
            try {
                dumper.dump(spreadsheet, new File(line));
            } catch (IOException e) {
                System.out.println("Saving failed: " + e.getMessage());
                continue;
            }
            
            break;
        }
        
        System.out.println("File saved successfully!");
    }

}
