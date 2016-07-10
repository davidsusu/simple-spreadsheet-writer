package hu.webarticum.simple_spreadsheet_writer.example;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.ObjectInputStream;

import hu.webarticum.simple_spreadsheet_writer.Spreadsheet;


public class UnserializeExample implements Example {

    @Override
    public String getLabel() {
        return "Unserialize an existing spreadsheet";
    }

    @Override
    public Spreadsheet create() {
        Spreadsheet spreadsheet = null;
        System.out.print("Input file: ");
        try {
            String line = (new BufferedReader(new InputStreamReader(System.in))).readLine();
            File file = new File(line);
            ObjectInputStream objectInputStream = new ObjectInputStream(new FileInputStream(file));
            spreadsheet = (Spreadsheet)objectInputStream.readObject();
            objectInputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (spreadsheet != null) {
            return spreadsheet;
        } else {
            return new Spreadsheet();
        }
    }

}
