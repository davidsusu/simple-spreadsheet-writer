package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.io.OutputStream;


public class SerializeSpreadsheetDumper implements SpreadsheetDumper {

    @Override
    public String getDefaultExtension() {
        return "obj";
    }

    @Override
    public void dump(Spreadsheet spreadsheet, File file) throws IOException {
        dump(spreadsheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadsheet, OutputStream outputStream) throws IOException {
        ObjectOutputStream objectOutputStream = null;
        objectOutputStream = new ObjectOutputStream(outputStream);
        objectOutputStream.writeObject(spreadsheet);
        objectOutputStream.close();
    }

}
