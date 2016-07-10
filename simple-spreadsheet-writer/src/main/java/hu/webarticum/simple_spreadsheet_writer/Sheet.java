package hu.webarticum.simple_spreadsheet_writer;

import java.nio.channels.UnsupportedAddressTypeException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.commons.collections4.IteratorUtils;

import hu.webarticum.simple_spreadsheet_writer.util.MergeSortedIterator;

public class Sheet implements Iterable<Sheet.CellEntry> {

    static public final int ITERATOR_CELLS = 1;
    static public final int ITERATOR_ROWS = 2;
    static public final int ITERATOR_COLUMNS = 4;
    static public final int ITERATOR_AREAS = 8;
    static public final int ITERATOR_COMBINED = 14;
    static public final int ITERATOR_MERGES = 16;
    static public final int ITERATOR_ALL = 32;
    static public final int ITERATOR_FULL = 64;
    
    private TreeMap<Integer, Row> rows = new TreeMap<Integer, Row>();

    private TreeMap<Integer, Column> columns = new TreeMap<Integer, Column>();
    
    final public List<Area> areas = new ArrayList<Area>();

    final public List<Range> merges = new ArrayList<Range>();
    
    public Sheet() {
    }
    
    public Sheet(Sheet baseSheet) {
        for (Map.Entry<Integer, Row> entry: baseSheet.rows.entrySet()) {
            Integer rowIndex = entry.getKey();
            Row baseRow = entry.getValue();
            Row row = new Row();
            row.height = baseRow.height;
            row.format = new Format(baseRow.format);
            for (Map.Entry<Integer, Cell> _entry: baseRow.cells.entrySet()) {
                Integer columnIndex = _entry.getKey();
                Cell baseCell = _entry.getValue();
                Cell cell = new Cell(baseCell);
                row.cells.put(columnIndex, cell);
            }
            this.rows.put(rowIndex, row);
        }
        for (Map.Entry<Integer, Column> entry: baseSheet.columns.entrySet()) {
            Integer columnIndex = entry.getKey();
            Column baseColumn = entry.getValue();
            Column column = new Column();
            column.width = baseColumn.width;
            column.format = new Format(baseColumn.format);
            this.columns.put(columnIndex, column);
        }
        for (Area baseArea: baseSheet.areas) {
            Area area = new Area();
            for (Range baseRange: baseArea.ranges) {
                area.ranges.add(new Range(baseRange));
            }
            area.format = new Format(baseArea.format);
            this.areas.add(area);
        }
    }
    
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
            if ((iteratorType & ITERATOR_MERGES) > 0) {
                positionIterators.add(new MergesPositionIterator());
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
        if (!merges.isEmpty()) {
            return false;
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
        for (Range range: merges) {
            int mergeMinRowIndex = Math.min(range.rowIndex1, range.rowIndex2);
            if (mergeMinRowIndex < minRowIndex) {
                minRowIndex = mergeMinRowIndex;
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
                if (areaMaxRowIndex > maxRowIndex) {
                    maxRowIndex = areaMaxRowIndex;
                }
            }
        }
        for (Range range: merges) {
            int mergeMaxRowIndex = Math.max(range.rowIndex1, range.rowIndex2);
            if (mergeMaxRowIndex > maxRowIndex) {
                maxRowIndex = mergeMaxRowIndex;
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
                if (areaMaxColumnIndex > maxColumnIndex) {
                    maxColumnIndex = areaMaxColumnIndex;
                }
            }
        }
        return maxColumnIndex;
    }

    public void write(int rowIndex, int ColumnIndex, String text) {
        setCell(rowIndex, ColumnIndex, new Cell(text));
    }

    public void write(int rowIndex, int ColumnIndex, String text, Format format) {
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
        insertColumn(columnIndex, null);
    }
    
    public void insertColumn(int columnIndex, Column column) {
        if (column == null) {
            insertIntoTreeMap(columns, columnIndex);
        } else {
            insertIntoTreeMap(columns, columnIndex, column);
        }
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            insertIntoTreeMap(entry.getValue().cells, columnIndex);
        }
        for (Area area: areas) {
            Iterator<Range> iterator = area.ranges.iterator();
            while (iterator.hasNext()) {
                Range range = iterator.next();
                insertColumnToRange(range, columnIndex);
            }
        }
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            insertColumnToRange(range, columnIndex);
        }
    }
    
    public void removeColumn(int columnIndex) {
        columns.remove(columnIndex);
    }

    public void cutColumn(int columnIndex) {
        cutFromTreeMap(columns, columnIndex);
        for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
            TreeMap<Integer, Cell> rowCells = entry.getValue().cells;
            cutFromTreeMap(rowCells, columnIndex);
        }
        for (Area area: areas) {
            Iterator<Range> iterator = area.ranges.iterator();
            while (iterator.hasNext()) {
                Range range = iterator.next();
                if (cutColumnFromRange(range, columnIndex)) {
                    iterator.remove();
                }
            }
        }
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            if (cutColumnFromRange(range, columnIndex)) {
                iterator.remove();
            }
        }
    }

    public TreeSet<Integer> getColumnIndexes() {
        return new TreeSet<Integer>(columns.keySet());
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
        insertRow(rowIndex, null);
    }
    
    public void insertRow(int rowIndex, Row row) {
        if (row == null) {
            insertIntoTreeMap(rows, rowIndex);
        } else {
            insertIntoTreeMap(rows, rowIndex, row);
        }
        for (Area area: areas) {
            Iterator<Range> iterator = area.ranges.iterator();
            while (iterator.hasNext()) {
                Range range = iterator.next();
                insertRowToRange(range, rowIndex);
            }
        }
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            insertRowToRange(range, rowIndex);
        }
    }
    
    public void removeRow(int rowIndex) {
        rows.remove(rowIndex);
    }

    public void cutRow(int rowIndex) {
        cutFromTreeMap(rows, rowIndex);
        for (Area area: areas) {
            Iterator<Range> iterator = area.ranges.iterator();
            while (iterator.hasNext()) {
                Range range = iterator.next();
                if (cutRowFromRange(range, rowIndex)) {
                    iterator.remove();
                }
            }
        }
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            if (cutRowFromRange(range, rowIndex)) {
                iterator.remove();
            }
        }
    }
    
    public TreeSet<Integer> getRowIndexes() {
        return new TreeSet<Integer>(rows.keySet());
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

    public Range getMerge(int rowIndex, int columnIndex) {
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            if (
                rowIndex == Math.min(range.rowIndex1, range.rowIndex2) &&
                columnIndex == Math.min(range.columnIndex1, range.columnIndex2)
            ) {
                return range;
            }
        }
        return null;
    }
    
    public boolean isHiddenByMerge(int rowIndex, int columnIndex) {
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            if (range.contains(rowIndex, columnIndex)) {
                if (
                    rowIndex != Math.min(range.rowIndex1, range.rowIndex2) ||
                    columnIndex != Math.min(range.columnIndex1, range.columnIndex2)
                ) {
                    return true;
                }
            }
        }
        return false;
    }
    
    public void removeMerges(int rowIndex, int columnIndex) {
        Iterator<Range> iterator = merges.iterator();
        while (iterator.hasNext()) {
            Range range = iterator.next();
            if (range.contains(rowIndex, columnIndex)) {
                iterator.remove();
            }
        }
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
    
    public void move(int verticalMove, int horizontalMove) {
        if (verticalMove == 0 && horizontalMove == 0) {
            return;
        }
        if (verticalMove != 0) {
            TreeMap<Integer, Row> newRows = new TreeMap<Integer, Row>();
            for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
                newRows.put(entry.getKey() + verticalMove, entry.getValue());
            }
            rows = newRows;
        }
        if (horizontalMove != 0) {
            TreeMap<Integer, Column> newColumns = new TreeMap<Integer, Column>();
            for (Map.Entry<Integer, Column> entry: columns.entrySet()) {
                newColumns.put(entry.getKey() + horizontalMove, entry.getValue());
            }
            columns = newColumns;
            for (Map.Entry<Integer, Row> entry: rows.entrySet()) {
                Row row = entry.getValue();
                TreeMap<Integer, Cell> newCells = new TreeMap<Integer, Cell>();
                for (Map.Entry<Integer, Cell> _entry: row.cells.entrySet()) {
                    newCells.put(_entry.getKey() + horizontalMove, _entry.getValue());
                }
                row.cells = newCells;
            }
        }
        for (Area area: areas) {
            for (Range range: area.ranges) {
                range.rowIndex1 += verticalMove;
                range.columnIndex1 += horizontalMove;
                range.rowIndex2 += verticalMove;
                range.columnIndex2 += horizontalMove;
            }
        }
        for (Range range: merges) {
            range.rowIndex1 += verticalMove;
            range.columnIndex1 += horizontalMove;
            range.rowIndex2 += verticalMove;
            range.columnIndex2 += horizontalMove;
        }
    }
    
    public int[] moveToNonNegative() {
        int verticalMove = 0;
        int horizontalMove = 0;
        if (!isEmpty()) {
            int minRowIndex = getMinRowIndex();
            if (minRowIndex < 0) {
                verticalMove = -minRowIndex;
            }
            int minColumnIndex = getMinColumnIndex();
            if (minColumnIndex < 0) {
                horizontalMove = -minColumnIndex;
            }
        }
        move(verticalMove, horizontalMove);
        return new int[]{verticalMove, horizontalMove};
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

    private boolean cutColumnFromRange(Range range, int columnIndex) {
        if (range.columnIndex1 == range.columnIndex2) {
            return (range.columnIndex1 == columnIndex);
        }
        boolean inverted = (range.columnIndex1 > range.columnIndex2);
        int min = Math.min(range.columnIndex1, range.columnIndex2);
        int max = Math.max(range.columnIndex1, range.columnIndex2);
        if (columnIndex >= min && columnIndex <= max) {
            if (inverted) {
                range.columnIndex1--;
            } else {
                range.columnIndex2--;
            }
        }
        return false;
    }

    private void insertColumnToRange(Range range, int columnIndex) {
        boolean inverted = (range.columnIndex1 > range.columnIndex2);
        int min = Math.min(range.columnIndex1, range.columnIndex2);
        int max = Math.max(range.columnIndex1, range.columnIndex2);
        if (columnIndex > min && columnIndex <= max) {
            if (inverted) {
                range.columnIndex1++;
            } else {
                range.columnIndex2++;
            }
        }
    }

    private boolean cutRowFromRange(Range range, int rowIndex) {
        if (range.rowIndex1 == range.rowIndex2) {
            return (range.rowIndex1 == rowIndex);
        }
        boolean inverted = (range.rowIndex1 > range.rowIndex2);
        int min = Math.min(range.rowIndex1, range.rowIndex2);
        int max = Math.max(range.rowIndex1, range.rowIndex2);
        if (rowIndex >= min && rowIndex <= max) {
            if (inverted) {
                range.rowIndex1--;
            } else {
                range.rowIndex2--;
            }
        }
        return false;
    }

    private void insertRowToRange(Range range, int rowIndex) {
        boolean inverted = (range.rowIndex1 > range.rowIndex2);
        int min = Math.min(range.rowIndex1, range.rowIndex2);
        int max = Math.max(range.rowIndex1, range.rowIndex2);
        if (rowIndex > min && rowIndex <= max) {
            if (inverted) {
                range.rowIndex1++;
            } else {
                range.rowIndex2++;
            }
        }
    }

    static public class Row {

        public int height = 0;
        
        public Format format = new Format();
        
        TreeMap<Integer, Cell> cells = new TreeMap<Integer, Cell>();
        
    }
    
    static public class Column {

        public int width = 0;
        
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

        public Range(Range baseRange) {
            this(baseRange.rowIndex1, baseRange.columnIndex1, baseRange.rowIndex2, baseRange.columnIndex2);
        }
        
        public boolean contains(int rowIndex, int columnIndex) {
            return (
                isBetween(rowIndex, rowIndex1, rowIndex2) &&
                isBetween(columnIndex, columnIndex1, columnIndex2)
            );
        }
        
        @Override
        public String toString() {
            return "Range{" + rowIndex1 + ", " + columnIndex1 + ", " + rowIndex1 + ", " + columnIndex2 + ", }";
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

        @Override
        public int hashCode() {
            return (rowIndex1 + rowIndex2 + columnIndex1 + columnIndex2) / 37;
        }
        
        @Override
        public boolean equals(Object other) {
            if (other instanceof Range) {
                Range otherRange = (Range)other;
                return (
                    this.rowIndex1 == otherRange.rowIndex1 &&
                    this.rowIndex2 == otherRange.rowIndex2 &&
                    this.columnIndex1 == otherRange.columnIndex1 &&
                    this.columnIndex2 == otherRange.columnIndex2
                );
            } else {
                return false;
            }
        }

    }
    
    static public class Area {
        
        public List<Range> ranges = new ArrayList<Range>();
        
        public Format format = new Format();
        
        public Area() {
        }
        
        public Area(int... rangeIndexes) {
            int size = rangeIndexes.length / 4;
            for (int i = 0; i < size; i++) {
                ranges.add(new Range(
                    rangeIndexes[i * 4],
                    rangeIndexes[i * 4 + 1],
                    rangeIndexes[i * 4 + 2],
                    rangeIndexes[i * 4 + 3]
                ));
            }
        }

        public Area(int[] rangeIndexes, Format format) {
            this(rangeIndexes);
            this.format = format;
        }

        public Area(
            int rowIndex1, int columnIndex1, int rowIndex2, int columnIndex2,
            Format format
        ) {
            this.ranges.add(new Range(rowIndex1, columnIndex1, rowIndex2, columnIndex2));
            this.format = format;
        }
        
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
                int rangeMaxRowIndex = Math.max(range.rowIndex1, range.rowIndex2);
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
                int rangeMaxColumnIndex = Math.max(range.columnIndex1, range.columnIndex2);
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

        public Cell(Cell baseCell) {
            this(baseCell.type, baseCell.text, new Format(baseCell.format));
        }

    }

    // XXX <String, FormatValue>?
    // TODO: rename to Style
    static public class Format extends HashMap<String, String> {

        private static final long serialVersionUID = 1L;
        
        public Format() {
        }
        
        public Format(Format baseFormat) {
            for (Map.Entry<String, String> entry: baseFormat.entrySet()) {
                this.put(entry.getKey(), entry.getValue());
            }
        }
        
        public Format(String... propertiesAndValues) {
            int size = propertiesAndValues.length / 2;
            for (int i = 0; i < size; i++) {
                this.put(propertiesAndValues[i * 2], propertiesAndValues[i * 2 + 1]);
            }
        }
        
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

        public FormatList() {
            super();
        }

        public FormatList(FormatList formatList) {
            super(formatList);
        }
        
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
        
        public Format getComputedFormat() {
            FormatList formatList = new FormatList(areaFormats);
            formatList.add(columnFormat);
            formatList.add(rowFormat);
            formatList.add(cell.format);
            return formatList.merge();
        }
        
    }
    
    static public class PositionComparator implements Comparator<int[]> {

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

    static public class AreaPositionIterator implements Iterator<int[]> {

        private MergeSortedIterator<int[]> mergeIterator;
        
        public AreaPositionIterator(Area area) {
            List<Iterator<int[]>> rangeIterators = new ArrayList<Iterator<int[]>>();
            for (Range range: area.ranges) {
                rangeIterators.add(new RangePositionIterator(range));
            }
            mergeIterator = new MergeSortedIterator<int[]>(rangeIterators, new PositionComparator());
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
    
    protected class MergesPositionIterator implements Iterator<int[]> {

        Iterator<Range> rangeIterator;
        
        public MergesPositionIterator() {
            rangeIterator = merges.iterator();
        }
        
        @Override
        public boolean hasNext() {
            return rangeIterator.hasNext();
        }

        @Override
        public int[] next() {
            Range range = rangeIterator.next();
            return new int[]{range.rowIndex1, range.columnIndex1};
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
        
        int currentRelativeRowIndex;

        Set<Integer> columnIndexes;
        
        Iterator<Integer> currentColumnIndexIterator;

        public ColumnsPositionIterator() {
            columnIndexes = columns.keySet();
            minRowIndex = getMinRowIndex();
            height = getDefinedHeight();
            currentRelativeRowIndex = height - 1;
            currentColumnIndexIterator = columnIndexes.iterator();
        }
        
        @Override
        public boolean hasNext() {
            return (currentRelativeRowIndex < (height - 1) || currentColumnIndexIterator.hasNext());
        }
        
        @Override
        public int[] next() {
            if (currentColumnIndexIterator.hasNext()) {
                int columnIndex = currentColumnIndexIterator.next();
                return new int[]{minRowIndex + currentRelativeRowIndex, columnIndex};
            } else if (currentRelativeRowIndex < (height - 1)) {
                currentRelativeRowIndex++;
                currentColumnIndexIterator = columnIndexes.iterator();
                return next();
            } else {
                throw new NoSuchElementException();
            }
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
    
    static public class RangePositionIterator implements Iterator<int[]> {

        int startRowIndex;

        int endRowIndex;

        int startColumnIndex;

        int endColumnIndex;
        
        Integer currentRowIndex = null;

        Integer currentColumnIndex = null;
        
        public RangePositionIterator(Range range) {
            this.startRowIndex = Math.min(range.rowIndex1, range.rowIndex2);
            this.endRowIndex = Math.max(range.rowIndex1, range.rowIndex2);
            this.startColumnIndex = Math.min(range.columnIndex1, range.columnIndex2);
            this.endColumnIndex = Math.max(range.columnIndex1, range.columnIndex2);
        }

        @Override
        public boolean hasNext() {
            if (currentRowIndex == null) {
                return true;
            }
            return (
                currentRowIndex < endRowIndex ||
                (currentRowIndex == endRowIndex && currentColumnIndex < endColumnIndex)
            );
        }

        @Override
        public int[] next() {
            if (currentRowIndex == null) {
                currentRowIndex = startRowIndex;
                currentColumnIndex = startColumnIndex;
                return new int[]{currentRowIndex, currentColumnIndex};
            } else if (currentRowIndex > endRowIndex) {
                throw new NoSuchElementException();
            } else {
                currentColumnIndex++;
                if (currentColumnIndex > endColumnIndex) {
                    currentColumnIndex = startColumnIndex;
                    currentRowIndex ++;
                    if (currentRowIndex > endRowIndex) {
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
