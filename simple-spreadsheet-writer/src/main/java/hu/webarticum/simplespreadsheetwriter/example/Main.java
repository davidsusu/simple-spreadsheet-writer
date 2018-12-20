package hu.webarticum.simplespreadsheetwriter.example;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import hu.webarticum.simplespreadsheetwriter.HssfApachePoiSpreadsheetDumper;
import hu.webarticum.simplespreadsheetwriter.HtmlSpreadsheetDumper;
import hu.webarticum.simplespreadsheetwriter.OdfToolkitSpreadsheetDumper;
import hu.webarticum.simplespreadsheetwriter.SerializeSpreadsheetDumper;
import hu.webarticum.simplespreadsheetwriter.Spreadsheet;
import hu.webarticum.simplespreadsheetwriter.SpreadsheetDumper;
import hu.webarticum.simplespreadsheetwriter.XssfApachePoiSpreadsheetDumper;

public class Main {

    public static void main(String[] args) {
        List<Example> examples = new ArrayList<Example>();
        examples.add(new TimeTableExample());
        examples.add(new LotOfCellsExample());
        examples.add(new UnserializeExample());
        
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
        dumpers.add(new SerializeSpreadsheetDumper());
        
        Map<String, SpreadsheetDumper> dumperMap = new LinkedHashMap<String, SpreadsheetDumper>();
        for (SpreadsheetDumper dumper: dumpers) {
            dumperMap.put(dumper.getDefaultExtension(), dumper);
        }
        
        String defaultType = "ods";
        
        File outFile = null;
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
            
            if (!line.contains("/")) {
                line = "/tmp/" + line;
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

            outFile = new File(line);
            
            SpreadsheetDumper dumper = dumperMap.get(type);
            
            if (spreadsheet == null) {
                spreadsheet = example.create();
            }

            try {
                dumper.dump(spreadsheet, outFile);
            } catch (IOException e) {
                System.out.println("Saving failed: " + e.getMessage());
                continue;
            }
            
            break;
        }
        
        System.out.println("File saved successfully!");
        
        Desktop desktop = null;
        try {
            desktop = Desktop.getDesktop();
        } catch (UnsupportedOperationException e) {
        }
        if (desktop == null) {
            System.out.println("No desktop is available!");
        }
        
        try {
            desktop.open(outFile);
        } catch (IOException e) {
            System.out.println("Failed to open file!");
        }
    }

}
