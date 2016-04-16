package hu.webarticum.simple_spreadsheet_writer;

import java.nio.channels.UnsupportedAddressTypeException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.commons.collections4.IteratorUtils;

import hu.webarticum.simple_spreadsheet_writer.util.MergeSortedIterator;

public class Sheet implements Iterable<Sheet.CellEntry> {

    static public final int ITERATOR_CELLS = 1;
    static public final int ITERATOR_AREAS = 2;
    static public final int ITERATOR_COLUMNS = 4;
    static public final int ITERATOR_ROWS = 8;
    static public final int ITERATOR_COMBINED = 14;
    static public final int ITERATOR_ALL = 16;
    static public final int ITERATOR_FULL = 32;
    
    private TreeMap<Integer, Row> rows = new TreeMap<Integer, Row>();

    private TreeMap<Integer, Column> columns = new TreeMap<Integer, Column>();
    
    private Set<Area> areas = new LinkedHashSet<Area>();

    @Override
    public Iterator<CellEntry> iterator() {
        return iterator(ITERATOR_FULL);
    }
    
    public Iterator<CellEntry> iterator(int iteratorType) {
        if (iteratorType == 0 || isEmpty()) {
            return IteratorUtils.<CellEntry>emptyIterator();
        } else if ((iteratorType & (ITERATOR_ALL | ITERATOR_FULL)) > 0) {
            return new PositionCellEntryIterator(new AllPositionIterator((iteratorType & ITERATOR_FULL) > 0));
        } else {
            List<Iterator<int[]>> positionIterators = new ArrayList<Iterator<int[]>>();
            if ((iteratorType & ITERATOR_CELLS) > 0) {
                positionIterators.add(new CellPositionIterator());
            }
            if ((iteratorType & ITERATOR_ROWS) > 0) {
                positionIterators.add(new RowsPositionIterator());
            }
            if ((iteratorType & ITERATOR_COLUMNS) > 0) {
                positionIterators.add(new ColumnsPositionIterator());
            }
            if ((iteratorType & ITERATOR_AREAS) > 0) {
                for (Area area: areas) {
                    if (!area.isEmpty()) {
                        positionIterators.add(new AreaPositionIterator(area));
                    }
                }
            }
            return new PositionCellEntryIterator(new MergeSortedIterator<int[]>(positionIterators, new PositionComparator()));
        }
    }
    
    public boolean isEmpty() {
        if (!rows.isEmpty()) {
            return false;
        }
        if (!columns.isEmpty()) {
            return false;
        }
        for (Area area: areas) {
            if (!area.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    public int getDefinedWidth() {
        if (isEmpty()) {
            return 0;
        } else {
            return getMaxColumnIndex() - getMinColumnIndex() + 1;
        }
    }
    
    public int getDefinedHeight() {
        if (isEmpty()) {
            return 0;
        } else {
            return getMaxRowIndex() - getMinRowIndex() + 1;
        }
    }

    public int getWidth() {
        if (isEmpty()) {
            return 0;
        } else {
            return getMaxColumnIndex() - Math.min(0, getMinColumnIndex()) + 1;
        }
    }
    
    public int getHeight() {
        if (isEmpty()) {
            return 0;
        } else {
            return getMaxRowIndex() - Math.min(0, getMinRowIndex()) + 1;
        }
    }
    
    public boolean hasNegative() {
        return (!isEmpty() && getMinRowIndex() < 0);
    }

    public int getMinRowIndex() {
        int minRowIndex;
        if (!rows.isEmpty()) {
            minRowIndex = rows.firstKey();
        } else {
            minRowIndex = Integer.MAX_VALUE;
        }
        for (Area area: areas) {
            if (!area.isEmpty()) {
                int areaMinRowIndex = area.getMinRowIndex();
                if (areaMinRowIndex < minRowIndex) {
                    minRowIndex = areaMinRowIndex;
                }
            }
        }
        return minRowIndex;
    }

    public int getMaxRowIndex() {
        int maxRowIndex;
        if (!rows.isEmpty()) {
            maxRowIndex = rows.lastKey();
        } else {
            maxRowIndex = Integer.MIN_VALUE;
        }
        for (Area area: areas) {
            if (!area.isEmpty()) {
                int areaMaxRowIndex = area.getMaxRowIndex();
                if (areaMaxRowIndex < maxRowIndex) {
                    maxRowIndex = areaMaxRowIndex;
                }
            }
        }
        return maxRowIndex;
    }

    public int getMinColumnIndex() {
        int minColumnIndex;
        if (!columns.isEmpty()) {
            minColumnIndex = columns.firstKey();
        } else {
            minColumnIndex = Integer.MAX_VALUE;
        }
        for (Map.Entry<Integer, Row> rowEntry: rows.entrySet()) {
            TreeMap<Integer, Cell> cells = rowEntry.getValue().cells;
            if (!cells.isEmpty()) {
                int rowMinColumnIndex = cells.firstKey();
                if (rowMinColumnIndex < minColumnIndex) {
                    minColumnIndex = rowMinColumnIndex;
                }
            }
        }
        for (Area area: areas) {
            if (!area.isEmpty()) {
                int areaMinColumnIndex = area.getMinColumnIndex();
                if (areaMinColumnIndex < minColumnIndex) {
                    minColumnIndex = areaMinColumnIndex;
                }
            }
        }
        return minColumnIndex;
    }

    public int getMaxColumnIndex() {
        int maxColumnIndex;
        if (!columns.isEmpty()) {
            maxColumnIndex = columns.lastKey();
        } else {
            maxColumnIndex = Integer.MIN_VALUE;
        }
        for (Map.Entry<Integer, Row> rowEntry: rows.entrySet()) {
            TreeMap<Integer, Cell> cells = rowEntry.getValue().cells;
            if (!cells.isEmpty()) {
                int rowMaxColumnIndex = cells.lastKey();
                if (rowMaxColumnIndex > maxColumnIndex) {
                    maxColumnIndex = rowMaxColumnIndex;
                }
            }
        }
        for (Area area: areas) {
            if (!area.isEmpty()) {
                int areaMaxColumnIndex = area.getMaxColumnIndex();
                if (areaMaxColumnIndex < maxColumnIndex) {
                    maxColumnIndex = areaMaxColumnIndex;
                }
            }
        }
        return maxColumnIndex;
    }

    public void writeText(int rowIndex, int ColumnIndex, String text) {
        setCell(rowIndex, ColumnIndex, new Cell(text));
    }

    public void writeText(int rowIndex, int ColumnIndex, String text, Format format) {
        setCell(rowIndex, ColumnIndex, new Cell(text, format));
    }
    
    public boolean hasColumn(int columnIndex) {
        return columns.containsKey(columnIndex);
    }
    
    public Column getColumn(int columnIndex) {
        if (columns.containsKey(columnIndex)) {
            return columns.get(columnIndex);
        } else {
            return null;
        }
    }

    public Column getColumn(int columnIndex, Column fallbackColumn) {
        Column column = getColumn(columnIndex);
        if (column==null && fallbackColumn!=null) {
            column = fallbackColumn;
            columns.put(columnIndex, fallbackColumn);
        }
        return column;
    }

    public void setColumn(int columnIndex, Column column) {
        if (column!=null) {
            columns.put(columnIndex, column);
        } else {
            columns.remove(columnIndex);
        }
    }

    public void insertColumn(int columnIndex) {
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            insertIntoTreeMap(entry.getValue().cells, columnIndex);
        }
        insertIntoTreeMap(columns, columnIndex);
    }
    
    public void insertColumn(int columnIndex, Column column) {
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            insertIntoTreeMap(entry.getValue().cells, columnIndex);
        }
        insertIntoTreeMap(columns, columnIndex, column);
    }
    
    public void removeColumn(int columnIndex) {
        columns.remove(columnIndex);
    }

    public void cutColumn(int columnIndex) {
        columns.remove(columnIndex);
        List<Integer> rowIndexesToRemove = new ArrayList<Integer>();
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            TreeMap<Integer, Cell> rowCells = entry.getValue().cells;
            cutFromTreeMap(rowCells, columnIndex);
            if (rowCells.isEmpty()) {
                rowIndexesToRemove.add(entry.getKey());
            }
        }
        for (Integer rowIndex: rowIndexesToRemove) {
            rows.remove(rowIndex);
        }
    }
    
    public boolean hasRow(int rowIndex) {
        return rows.containsKey(rowIndex);
    }
    
    public Row getRow(int rowIndex) {
        if (rows.containsKey(rowIndex)) {
            return rows.get(rowIndex);
        } else {
            return null;
        }
    }

    public Row getRow(int rowIndex, Row fallbackRow) {
        Row row = getRow(rowIndex);
        if (row==null && fallbackRow!=null) {
            row = fallbackRow;
            rows.put(rowIndex, fallbackRow);
        }
        return row;
    }

    public void setRow(int rowIndex, Row row) {
        if (row!=null) {
            rows.put(rowIndex, row);
        } else {
            rows.remove(rowIndex);
        }
    }

    public void insertRow(int rowIndex) {
        insertIntoTreeMap(rows, rowIndex);
    }
    
    public void insertRow(int rowIndex, Row row) {
        insertIntoTreeMap(rows, rowIndex, row);
    }
    
    public void removeRow(int rowIndex) {
        rows.remove(rowIndex);
    }

    public void cutRow(int rowIndex) {
        cutFromTreeMap(rows, rowIndex);
    }
    
    public boolean hasCell(int rowIndex, int columnIndex) {
        if (!rows.containsKey(rowIndex)) {
            return false;
        }
        return rows.get(rowIndex).cells.containsKey(columnIndex);
    }
    
    public Cell getCell(int rowIndex, int columnIndex) {
        if (!rows.containsKey(rowIndex)) {
            return null;
        }
        Row row = rows.get(rowIndex);
        if (!row.cells.containsKey(columnIndex)) {
            return null;
        }
        return row.cells.get(columnIndex);
    }

    public Cell getCell(int rowIndex, int columnIndex, Cell fallbackCell) {
        Row row;
        if (rows.containsKey(rowIndex)) {
            row = rows.get(rowIndex);
        } else if (fallbackCell==null) {
            return null;
        } else {
            row = new Row();
            rows.put(rowIndex, row);
        }
        Cell cell;
        if (row.cells.containsKey(columnIndex)) {
            cell = row.cells.get(columnIndex);
        } else if (fallbackCell==null) {
            return null;
        } else {
            cell = fallbackCell;
            row.cells.put(columnIndex, cell);
        }
        return cell;
    }

    public TreeMap<Integer, Cell> getColumnCells(int columnIndex) {
        TreeMap<Integer, Cell> cells = new TreeMap<Integer, Cell>();
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            TreeMap<Integer, Cell> rowCells = entry.getValue().cells;
            if (rowCells.containsKey(columnIndex)) {
                cells.put(entry.getKey(), rowCells.get(columnIndex));
            }
        }
        return cells;
    }
    
    public TreeMap<Integer, Cell> getRowCells(int rowIndex) {
        if (rows.containsKey(rowIndex)) {
            return new TreeMap<Integer, Cell>(rows.get(rowIndex).cells);
        } else {
            return new TreeMap<Integer, Cell>();
        }
    }
    
    public void setCell(int rowIndex, int columnIndex, Cell cell) {
        Row row;
        if (rows.containsKey(rowIndex)) {
            row = rows.get(rowIndex);
        } else if (cell==null) {
            return;
        } else {
            row = new Row();
            rows.put(rowIndex, row);
        }
        if (cell!=null) {
            row.cells.put(columnIndex, cell);
        } else {
            row.cells.remove(columnIndex);
            if (row.cells.isEmpty()) {
                rows.remove(rowIndex);
            }
        }
    }

    public void removeCell(int rowIndex, int columnIndex) {
        setCell(rowIndex, columnIndex, null);
    }
    
    public void addArea(Area area) {
        areas.add(area);
    }

    public void removeArea(Area area) {
        areas.remove(area);
    }
    
    public CellEntry getCellEntry(int rowIndex, int columnIndex) {
        CellEntry cellEntry = new CellEntry();
        cellEntry.rowIndex = rowIndex;
        cellEntry.columnIndex = columnIndex;
        if (rows.containsKey(rowIndex)) {
            Row row = rows.get(rowIndex);
            cellEntry.rowFormat = row.format;
            if (row.cells.containsKey(columnIndex)) {
                cellEntry.cell = row.cells.get(columnIndex);
                cellEntry.exists = true;
            } else {
                cellEntry.cell = new Cell();
                cellEntry.exists = false;
            }
        } else {
            cellEntry.cell = new Cell();
            cellEntry.rowFormat = new Format();
            cellEntry.exists = false;
        }
        if (columns.containsKey(columnIndex)) {
            Column column = columns.get(columnIndex);
            cellEntry.columnFormat = column.format;
        } else {
            cellEntry.columnFormat = new Format();
        }
        FormatList areaFormat = new FormatList();
        for (Area area: areas) {
            if (area.contains(rowIndex, columnIndex)) {
                areaFormat.add(area.format);
            }
        }
        cellEntry.areaFormats = areaFormat;
        return cellEntry;
    }
    
    private <T> void cutFromTreeMap(TreeMap<Integer, T> map, int index) {
        map.remove(index);
        SortedMap<Integer, T> afterItemsView = map.subMap(index+1, Integer.MAX_VALUE);
        TreeMap<Integer, T> afterItems = new TreeMap<Integer, T>(afterItemsView);
        afterItemsView.clear();
        for (Map.Entry<Integer, T> entry: afterItems.entrySet()) {
            map.put(entry.getKey()-1, entry.getValue());
        }
    }

    private <T> void insertIntoTreeMap(TreeMap<Integer, T> map, int index) {
        SortedMap<Integer, T> afterItemsView = map.subMap(index, Integer.MAX_VALUE);
        TreeMap<Integer, T> afterItems = new TreeMap<Integer, T>(afterItemsView);
        afterItemsView.clear();
        for (Map.Entry<Integer, T> entry: afterItems.entrySet()) {
            map.put(entry.getKey()+1, entry.getValue());
        }
    }

    private <T> void insertIntoTreeMap(TreeMap<Integer, T> map, int index, T item) {
        insertIntoTreeMap(map, index);
        map.put(index, item);
    }
    
    static public class Row {

        public Integer width = null;
        
        public Format format = new Format();
        
        TreeMap<Integer, Cell> cells = new TreeMap<Integer, Cell>();
        
    }
    
    static public class Column {

        public Integer width = null;
        
        public Format format = new Format();
        
    }

    static public class Range {
        
        public int rowIndex1;
        
        public int columnIndex1;
        
        public int rowIndex2;
        
        public int columnIndex2;
        
        public Range(int rowIndex1, int columnIndex1, int rowIndex2, int columnIndex2) {
            this.rowIndex1 = rowIndex1;
            this.columnIndex1 = columnIndex1;
            this.rowIndex2 = rowIndex2;
            this.columnIndex2 = columnIndex2;
        }

        public boolean contains(int rowIndex, int columnIndex) {
            return (
                isBetween(rowIndex, rowIndex1, rowIndex2) &&
                isBetween(columnIndex, columnIndex1, columnIndex2)
            );
        }

        private boolean isBetween(int number, int bound1, int bound2) {
            if (bound1<bound2) {
                return (number>=bound1 && number<=bound2);
            } else if (bound1>bound2) {
                return (number>=bound2 && number<=bound1);
            } else {
                return (number==bound1);
            }
        }

    }
    
    static public class Area {
        
        public List<Range> ranges = new ArrayList<Range>();
        
        public Format format = new Format();
        
        public boolean contains(int rowIndex, int columnIndex) {
            for (Range range: ranges) {
                if (range.contains(rowIndex, columnIndex)) {
                    return true;
                }
            }
            return false;
        }
        
        public boolean isEmpty() {
            return ranges.isEmpty();
        }
        
        public int getMinRowIndex() {
            int minRowIndex = Integer.MAX_VALUE;
            for (Range range: ranges) {
                int rangeMinRowIndex = Math.min(range.rowIndex1, range.rowIndex2);
                if (rangeMinRowIndex < minRowIndex) {
                    minRowIndex = rangeMinRowIndex;
                }
            }
            return minRowIndex;
        }

        public int getMaxRowIndex() {
            int maxRowIndex = Integer.MIN_VALUE;
            for (Range range: ranges) {
                int rangeMaxRowIndex = Math.min(range.rowIndex1, range.rowIndex2);
                if (rangeMaxRowIndex > maxRowIndex) {
                    maxRowIndex = rangeMaxRowIndex;
                }
            }
            return maxRowIndex;
        }

        public int getMinColumnIndex() {
            int minColumnIndex = Integer.MAX_VALUE;
            for (Range range: ranges) {
                int rangeMinColumnIndex = Math.min(range.columnIndex1, range.columnIndex2);
                if (rangeMinColumnIndex < minColumnIndex) {
                    minColumnIndex = rangeMinColumnIndex;
                }
            }
            return minColumnIndex;
        }

        public int getMaxColumnIndex() {
            int maxColumnIndex = Integer.MIN_VALUE;
            for (Range range: ranges) {
                int rangeMaxColumnIndex = Math.min(range.columnIndex1, range.columnIndex2);
                if (rangeMaxColumnIndex > maxColumnIndex) {
                    maxColumnIndex = rangeMaxColumnIndex;
                }
            }
            return maxColumnIndex;
        }

    }
    
    static public class Cell {

        public enum TYPE {TEXT};
        
        public String text = "";

        public TYPE type = TYPE.TEXT;
        
        public Format format = new Format();
        
        public Cell() {
            this(TYPE.TEXT, "", new Format());
        }

        public Cell(String text) {
            this(TYPE.TEXT, text, new Format());
        }

        public Cell(TYPE type, String text) {
            this(type, text, new Format());
        }

        public Cell(String text, Format format) {
            this(TYPE.TEXT, text, format);
        }

        public Cell(TYPE type, String text, Format format) {
            this.type = type;
            this.text = text;
            this.format = format;
        }
        
    }

    // XXX <String, FormatValue>?
    static public class Format extends HashMap<String, String> {

        private static final long serialVersionUID = 1L;
        
        public String toCssString() {
            StringBuilder cssTextBuilder = new StringBuilder();
            for (Map.Entry<String, String> entry: entrySet()) {
                cssTextBuilder.append(escapeCssProperty(entry.getKey()) + ":" + escapeCssValue(entry.getValue()) + ";");
            }
            return cssTextBuilder.toString();
        }

        private String escapeCssProperty(String name) {
            // TODO
            return name;
        }

        private String escapeCssValue(String value) {
            // TODO
            return value;
        }
        
    }
    
    static public class FormatList extends ArrayList<Format> {

        private static final long serialVersionUID = 1L;
        
        public Format merge() {
            Format result = new Format();
            for (Format format: this) {
                for (Map.Entry<String, String> entry: format.entrySet()) {
                    result.put(entry.getKey(), entry.getValue());
                }
            }
            return result;
        }
        
    }
    
    public class CellEntry {
        
        public boolean exists;
        
        public int rowIndex;
        
        public int columnIndex;
        
        public Cell cell;

        public Format rowFormat;

        public Format columnFormat;
        
        public FormatList areaFormats;
        
    }
    
    protected class PositionComparator implements Comparator<int[]> {

        @Override
        public int compare(int[] position1, int[] position2) {
            int result = Integer.compare(position1[0], position2[0]);
            if (result == 0) {
                result = Integer.compare(position1[1], position2[1]);
            }
            return result;
        }
        
    }
    
    protected class CellPositionIterator implements Iterator<int[]> {
        
        int currentRowIndex;
        
        int currentColumnIndex;
        
        Iterator<Map.Entry<Integer, Row>> rowEntryIterator;
        
        Iterator<Integer> rowKeyIterator;
        
        public CellPositionIterator() {
            rowEntryIterator = rows.entrySet().iterator();
            if (rows.isEmpty()) {
                rowKeyIterator = IteratorUtils.<Integer>emptyIterator();
            } else {
                nextRow();               
            }
        }
        
        @Override
        public boolean hasNext() {
            return (rowEntryIterator.hasNext()||rowKeyIterator.hasNext());
        }

        @Override
        public int[] next() {
            if (rowKeyIterator.hasNext()) {
                int columnIndex = rowKeyIterator.next();
                return new int[]{currentRowIndex, columnIndex};
            } else if (rowEntryIterator.hasNext()) {
                nextRow();
                return next();
            } else {
                throw new NoSuchElementException();
            }
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException();
        }
        
        private void nextRow() {
            Map.Entry<Integer, Row> rowEntry = rowEntryIterator.next();
            currentRowIndex = rowEntry.getKey();
            rowKeyIterator = rowEntry.getValue().cells.keySet().iterator();
        }
        
    }

    protected class AreaPositionIterator implements Iterator<int[]> {

        private MergeSortedIterator<int[]> mergeIterator;
        
        public AreaPositionIterator(Area area) {
            List<Iterator<int[]>> rangeIterators = new ArrayList<Iterator<int[]>>();
            for (Range range: area.ranges) {
                rangeIterators.add(new RangePositionIterator(range));
            }
            mergeIterator = new MergeSortedIterator<int[]>(rangeIterators);
        }
        
        @Override
        public boolean hasNext() {
            return mergeIterator.hasNext();
        }

        @Override
        public int[] next() {
            return mergeIterator.next();
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException();
        }
        
    }

    protected class RowsPositionIterator implements Iterator<int[]> {
        
        int minColumnIndex;
        
        int width;
        
        Iterator<Integer> rowIndexIterator;

        Integer currentRowIndex = null;
        
        int currentRelativeColumnIndex;
        
        public RowsPositionIterator() {
            rowIndexIterator = rows.keySet().iterator();
            minColumnIndex = getMinColumnIndex();
            width = getDefinedWidth();
            currentRelativeColumnIndex = width - 1;
        }
        
        @Override
        public boolean hasNext() {
            if (width == 0) {
                return false;
            } else if (currentRelativeColumnIndex < width - 1) {
                return true;
            } else {
                return rowIndexIterator.hasNext();
            }
        }
        
        @Override
        public int[] next() {
            if (width > 0 && (currentRelativeColumnIndex < width - 1 || rowIndexIterator.hasNext())) {
                currentRelativeColumnIndex++;
                if (currentRelativeColumnIndex == width) {
                    currentRelativeColumnIndex = 0;
                    currentRowIndex = rowIndexIterator.next();
                }
                return new int[]{currentRowIndex, minColumnIndex + currentRelativeColumnIndex};
            }
            throw new NoSuchElementException();
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException();
        }
        
    }

    protected class ColumnsPositionIterator implements Iterator<int[]> {
        
        int minRowIndex;
        
        int height;
        
        Iterator<Integer> columnIndexIterator;

        Integer currentColumnIndex = null;
        
        int currentRelativeRowIndex;
        
        public ColumnsPositionIterator() {
            columnIndexIterator = columns.keySet().iterator();
            minRowIndex = getMinRowIndex();
            height = getDefinedHeight();
            currentRelativeRowIndex = height - 1;
        }
        
        @Override
        public boolean hasNext() {
            if (height == 0) {
                return false;
            } else if (currentRelativeRowIndex < height - 1) {
                return true;
            } else {
                return columnIndexIterator.hasNext();
            }
        }
        
        @Override
        public int[] next() {
            if (height > 0 && (currentRelativeRowIndex < height - 1 || columnIndexIterator.hasNext())) {
                currentRelativeRowIndex++;
                if (currentRelativeRowIndex == height) {
                    currentRelativeRowIndex = 0;
                    currentColumnIndex = columnIndexIterator.next();
                }
                return new int[]{minRowIndex + currentRelativeRowIndex, currentColumnIndex};
            }
            throw new NoSuchElementException();
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException();
        }
        
    }
    
    protected class AllPositionIterator implements Iterator<int[]> {
        
        int width;
        
        int height;
        
        int minColumnIndex;
        
        int minRowIndex;

        int currentRelativeColumnIndex;

        int currentRelativeRowIndex;
        
        public AllPositionIterator(boolean full) {
            minColumnIndex = getMinColumnIndex();
            width = getDefinedWidth();
            if (full && width > 0 && minColumnIndex > 0) {
                width += minColumnIndex;
                minColumnIndex = 0;
            }
            currentRelativeColumnIndex = width - 1;
            minRowIndex = getMinRowIndex();
            height = getDefinedHeight();
            if (full && height > 0 && minRowIndex > 0) {
                height += minRowIndex;
                minRowIndex = 0;
            }
            currentRelativeRowIndex = -1;
        }
        
        @Override
        public boolean hasNext() {
            if (width == 0 || height == 0) {
                return false;
            } else if (currentRelativeRowIndex < height - 1) {
                return true;
            } else {
                return currentRelativeColumnIndex < width - 1;
            }
        }
        
        @Override
        public int[] next() {
            if (!hasNext()) {
                throw new NoSuchElementException();
            }
            currentRelativeColumnIndex++;
            if (currentRelativeColumnIndex == width) {
                currentRelativeColumnIndex = 0;
                currentRelativeRowIndex++;
            }
            return new int[]{minRowIndex + currentRelativeRowIndex, minColumnIndex + currentRelativeColumnIndex};
        }
        
        @Override
        public void remove() {
            throw new UnsupportedAddressTypeException();
        }
        
    }
    
    protected class RangePositionIterator implements Iterator<int[]> {

        int startRowIndex;

        int endRowIndex;

        int startColumnIndex;

        int endColumnIndex;
        
        Integer currentRowIndex = null;

        Integer currentColumnIndex = null;
        
        public RangePositionIterator(Range range) {
            this.startRowIndex = Math.min(range.rowIndex1, range.rowIndex2);
            this.endRowIndex = Math.max(range.rowIndex1, range.rowIndex2);
            this.startColumnIndex = Math.min(range.columnIndex1, range.rowIndex2);
            this.endColumnIndex = Math.max(range.columnIndex1, range.rowIndex2);
        }

        @Override
        public boolean hasNext() {
            if (currentRowIndex == null) {
                return true;
            }
            return (currentRowIndex < endRowIndex && currentColumnIndex < endColumnIndex);
        }

        @Override
        public int[] next() {
            if (currentRowIndex == null) {
                currentRowIndex = startRowIndex;
                currentColumnIndex = startColumnIndex;
                return new int[]{currentRowIndex, currentColumnIndex};
            } else if (currentRowIndex >= endRowIndex) {
                throw new NoSuchElementException();
            } else {
                currentColumnIndex++;
                if (currentColumnIndex > endColumnIndex) {
                    currentColumnIndex = startColumnIndex;
                    currentRowIndex ++;
                    if (currentRowIndex >= endRowIndex) {
                        throw new NoSuchElementException();
                    }
                }
                return new int[]{currentRowIndex, currentColumnIndex};
            }
        }

        @Override
        public void remove() {
            throw new UnsupportedOperationException();
        }
        
    }
    
    public class PositionCellEntryIterator implements Iterator<CellEntry> {

        private Iterator<int[]> positionIterator;
        
        PositionCellEntryIterator(Iterator<int[]> positionIterator) {
            this.positionIterator = positionIterator;
        }
        
        @Override
        public boolean hasNext() {
            return positionIterator.hasNext();
        }

        @Override
        public CellEntry next() {
            if (!positionIterator.hasNext()) {
                throw new NoSuchElementException();
            }
            int[] position = positionIterator.next();
            return getCellEntry(position[0], position[1]);
        }

        @Override
        public void remove() {
            positionIterator.remove();
        }
        
    }
        
}