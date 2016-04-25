package hu.webarticum.simple_spreadsheet_writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;


public class HtmlSpreadsheetDumper implements SpreadsheetDumper {

    @Override
    public String getDefaultExtension() {
        return "html";
    }
    
    @Override
    public void dump(Spreadsheet spreadsheet, File file) throws IOException {
        dump(spreadsheet, new FileOutputStream(file));
    }

    @Override
    public void dump(Spreadsheet spreadsheet, OutputStream outputStream) throws IOException {
        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.append("<!doctype html>\n");
        htmlBuilder.append("<html>\n");
        htmlBuilder.append("<head>\n");
        htmlBuilder.append("<meta http-equiv=\"Content-type\" content=\"text/html; charset=utf-8\">\n");
        htmlBuilder.append("<title>Spreadsheet</title>\n");
        htmlBuilder.append("</head>\n");
        htmlBuilder.append("<body>\n");
        htmlBuilder.append("<h1>Spreadsheet</h1>\n");
        for (Spreadsheet.Page page: spreadsheet) {
            Sheet sheet = page.sheet;
            if (sheet.hasNegative()) {
                sheet = new Sheet(sheet);
                sheet.moveToNonNegative();
            }
            int width = sheet.getWidth();
            htmlBuilder.append("<h2>" + page.label + "</h2>\n");
            htmlBuilder.append("<table style=\"border-collapse:collapse;table-layout:fixed;empty-cells:show;\">\n");
            htmlBuilder.append("<colgroup>\n");
            for (int columnIndex = 0; columnIndex < width; columnIndex++) {
                htmlBuilder.append("<col");
                if (sheet.hasColumn(columnIndex)) {
                    Sheet.Column column = sheet.getColumn(columnIndex);
                    if (column.width > 0) {
                        htmlBuilder.append(" style=\"width:" + column.width + "mm;\"");
                    }
                }
                htmlBuilder.append(" />\n");
            }
            htmlBuilder.append("</colgroup>\n");
            htmlBuilder.append("<tbody>\n");
            Iterator<Sheet.CellEntry> iterator = sheet.iterator(Sheet.ITERATOR_FULL);
            Integer previousRowIndex = null;
            while (iterator.hasNext()) {
                Sheet.CellEntry cellEntry = iterator.next();
                if (previousRowIndex == null || cellEntry.rowIndex != previousRowIndex) {
                    if (previousRowIndex != null) {
                        htmlBuilder.append("</tr>\n");
                    }
                    Sheet.Row row = sheet.getRow(cellEntry.rowIndex);
                    if (row.height > 0) {
                        htmlBuilder.append("<tr style=\"height:" + row.height + "mm;\">\n");
                    } else {
                        htmlBuilder.append("<tr>\n");
                    }
                }
                if (!sheet.isHiddenByMerge(cellEntry.rowIndex, cellEntry.columnIndex)) {
                    htmlBuilder.append("<td");
                    String cssText = cellEntry.getComputedFormat().toCssString();
                    if (!cssText.isEmpty()) {
                        htmlBuilder.append(" style=\"" + escape(cssText) + "\"");
                    }
                    Sheet.Range merge = sheet.getMerge(cellEntry.rowIndex, cellEntry.columnIndex);
                    if (merge != null) {
                        int mergeHorizontalSize = Math.abs(merge.columnIndex1 - merge.columnIndex2) + 1;
                        if (mergeHorizontalSize > 1) {
                            htmlBuilder.append(" colspan=\"" + mergeHorizontalSize + "\"");
                        }
                        int mergeVerticalSize = Math.abs(merge.rowIndex1 - merge.rowIndex2) + 1;
                        if (mergeVerticalSize > 1) {
                            htmlBuilder.append(" rowspan=\"" + mergeVerticalSize + "\"");
                        }
                    }
                    htmlBuilder.append(">");
                    htmlBuilder.append(escape(cellEntry.cell.text));
                    htmlBuilder.append("</td>\n");
                }
                previousRowIndex = cellEntry.rowIndex;
            }
            if (previousRowIndex != null) {
                htmlBuilder.append("</tr>\n");
            }
            htmlBuilder.append("</tbody>\n");
            htmlBuilder.append("</table>\n");
        }
        htmlBuilder.append("</body>\n");
        htmlBuilder.append("</html>\n");
        outputStream.write(htmlBuilder.toString().getBytes());
        outputStream.close();
    }
    
    protected String escape(String text) {
        text = text.replaceAll("&", "&amp;");
        text = text.replaceAll(">", "&gt;");
        text = text.replaceAll(">", "&lt;");
        text = text.replaceAll("\"", "&quot;");
        return text;
    }

}
