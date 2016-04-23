package hu.webarticum.simple_spreadsheet_writer;

import org.apache.commons.collections4.IteratorUtils;
import org.junit.Test;

import static org.junit.Assert.*;

import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class SheetTest {

    @Test
    public void testEmptySheet() {
        Sheet sheet = new Sheet();
        
        assertFalse(sheet.hasNegative());
        assertEquals(0, sheet.getWidth());
        assertEquals(0, sheet.getHeight());
        assertEquals(0, sheet.getDefinedWidth());
        assertEquals(0, sheet.getDefinedHeight());
        assertFalse(sheet.iterator().hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_CELLS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_ROWS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_COLUMNS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_AREAS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_COMBINED).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_ALL).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_FULL).hasNext());

        sheet.write(-1, -1, "Test text");

        assertTrue(sheet.hasNegative());
        assertEquals(1, sheet.getWidth());
        assertEquals(1, sheet.getHeight());

        sheet.removeCell(-1, -1);
        sheet.write(1, 1, "Test text");

        assertFalse(sheet.hasNegative());
        assertEquals(2, sheet.getWidth());
        assertEquals(2, sheet.getHeight());
        assertEquals(1, sheet.getDefinedWidth());
        assertEquals(1, sheet.getDefinedHeight());
        assertTrue(sheet.iterator().hasNext());
        assertTrue(sheet.iterator(Sheet.ITERATOR_CELLS).hasNext());
        assertTrue(sheet.iterator(Sheet.ITERATOR_ROWS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_COLUMNS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_AREAS).hasNext());
        assertTrue(sheet.iterator(Sheet.ITERATOR_COMBINED).hasNext());
        assertTrue(sheet.iterator(Sheet.ITERATOR_ALL).hasNext());
        assertTrue(sheet.iterator(Sheet.ITERATOR_FULL).hasNext());
        
        sheet.removeCell(1, 1);

        assertFalse(sheet.hasNegative());
        assertEquals(0, sheet.getWidth());
        assertEquals(0, sheet.getHeight());
        assertEquals(0, sheet.getDefinedWidth());
        assertEquals(0, sheet.getDefinedHeight());
        assertFalse(sheet.iterator().hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_CELLS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_ROWS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_COLUMNS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_AREAS).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_COMBINED).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_ALL).hasNext());
        assertFalse(sheet.iterator(Sheet.ITERATOR_FULL).hasNext());
    }

    @Test
    public void testNegativePositionSheet() {
        Sheet sheet = new Sheet();
        sheet.write(-3, -4, "Test text 1");
        sheet.write(1, 3, "Test text 2");
        assertTrue(sheet.hasNegative());
        assertEquals(8, sheet.getWidth());
        assertEquals(5, sheet.getHeight());
        assertEquals(8, sheet.getDefinedWidth());
        assertEquals(5, sheet.getDefinedHeight());
        assertTrue(sheet.hasCell(-3, -4));
        assertFalse(sheet.hasCell(0, 0));
        assertFalse(sheet.hasCell(1, 2));
        assertTrue(sheet.hasCell(1, 3));
    }
    
    @Test
    public void testSimpleSheet() {
        {
            Sheet sheet = new Sheet();
            sheet.write(1, 2, "Test text 1");
            sheet.write(2, 3, "Test text 2");
            assertFalse(sheet.hasNegative());
            assertEquals(4, sheet.getWidth());
            assertEquals(3, sheet.getHeight());
            assertEquals(2, sheet.getDefinedWidth());
            assertEquals(2, sheet.getDefinedHeight());
            assertFalse(sheet.hasCell(-3, -4));
            assertFalse(sheet.hasCell(0, 0));
            assertTrue(sheet.hasCell(1, 2));
            assertTrue(sheet.hasCell(2, 3));
            assertFalse(sheet.hasCell(2, 2));
            assertNull(sheet.getCell(3, 3));
            
            assertEquals("Fake text", sheet.getCell(3, 3, new Sheet.Cell("Fake text")).text);

            assertTrue(sheet.hasCell(3, 3));
            sheet.removeCell(3, 3);
            assertFalse(sheet.hasCell(3, 3));
            
            assertTrue(sheet.iterator().hasNext());
            assertFalse(sheet.iterator(0).hasNext());
            assertTrue(sheet.iterator(Sheet.ITERATOR_CELLS).hasNext());
            assertTrue(sheet.iterator(Sheet.ITERATOR_ROWS).hasNext());
            assertFalse(sheet.iterator(Sheet.ITERATOR_COLUMNS).hasNext());
            assertFalse(sheet.iterator(Sheet.ITERATOR_AREAS).hasNext());
            assertTrue(sheet.iterator(Sheet.ITERATOR_COMBINED).hasNext());
            assertTrue(sheet.iterator(Sheet.ITERATOR_ALL).hasNext());
            assertTrue(sheet.iterator(Sheet.ITERATOR_FULL).hasNext());
            assertEquals(2, sheet.getMinColumnIndex());
            assertEquals(3, sheet.getMaxColumnIndex());
            assertEquals(2, sheet.getDefinedWidth());
            assertEquals(1, sheet.getMinRowIndex());
            assertEquals(2, sheet.getMaxRowIndex());
            assertEquals(2, sheet.getDefinedHeight());
            {
                List<Sheet.CellEntry> list = IteratorUtils.toList(sheet.iterator(Sheet.ITERATOR_CELLS));
                assertEquals(2, list.size());
                {
                    Sheet.CellEntry cellEntry = list.get(0);
                    assertTrue(cellEntry.exists);
                    assertEquals(1, cellEntry.rowIndex);
                    assertEquals(2, cellEntry.columnIndex);
                    assertEquals("Test text 1", cellEntry.cell.text);
                }
                {
                    Sheet.CellEntry cellEntry = list.get(1);
                    assertTrue(cellEntry.exists);
                    assertEquals(2, cellEntry.rowIndex);
                    assertEquals(3, cellEntry.columnIndex);
                    assertEquals("Test text 2", cellEntry.cell.text);
                }
            }
            int[] boundIteratorTypes = new int[]{Sheet.ITERATOR_ROWS, Sheet.ITERATOR_COMBINED, Sheet.ITERATOR_ALL};
            for (int boundIteratorType: boundIteratorTypes) {
                List<Sheet.CellEntry> list = IteratorUtils.toList(sheet.iterator(boundIteratorType));
                assertEquals(4, list.size());
                {
                    Sheet.CellEntry cellEntry = list.get(0);
                    assertTrue(cellEntry.exists);
                    assertEquals(1, cellEntry.rowIndex);
                    assertEquals(2, cellEntry.columnIndex);
                    assertEquals("Test text 1", cellEntry.cell.text);
                }
                {
                    Sheet.CellEntry cellEntry = list.get(1);
                    assertFalse(cellEntry.exists);
                }
                {
                    Sheet.CellEntry cellEntry = list.get(2);
                    assertFalse(cellEntry.exists);
                }
                {
                    Sheet.CellEntry cellEntry = list.get(3);
                    assertTrue(cellEntry.exists);
                    assertEquals(2, cellEntry.rowIndex);
                    assertEquals(3, cellEntry.columnIndex);
                    assertEquals("Test text 2", cellEntry.cell.text);
                }
            }
            {
                List<Sheet.CellEntry> list = IteratorUtils.toList(sheet.iterator(Sheet.ITERATOR_COLUMNS));
                assertEquals(0, list.size());
            }
            {
                List<Sheet.CellEntry> list = IteratorUtils.toList(sheet.iterator(Sheet.ITERATOR_AREAS));
                assertEquals(0, list.size());
            }
            {
                List<Sheet.CellEntry> list = IteratorUtils.toList(sheet.iterator(Sheet.ITERATOR_FULL));
                assertEquals(12, list.size());
                {
                    Sheet.CellEntry cellEntry = list.get(6);
                    assertTrue(cellEntry.exists);
                    assertEquals(1, cellEntry.rowIndex);
                    assertEquals(2, cellEntry.columnIndex);
                    assertEquals("Test text 1", cellEntry.cell.text);
                }
                {
                    Sheet.CellEntry cellEntry = list.get(11);
                    assertTrue(cellEntry.exists);
                    assertEquals(2, cellEntry.rowIndex);
                    assertEquals(3, cellEntry.columnIndex);
                    assertEquals("Test text 2", cellEntry.cell.text);
                }
                for (int i = 0; i < 12; i++) {
                    if (i != 6 && i != 11) {
                        Sheet.CellEntry cellEntry = list.get(i);
                        assertFalse(cellEntry.exists);
                    }
                }
            }
        }
    }

    @Test
    public void testComplexSheet() {
        Sheet baseSheet = createComplexSheet();
        
        {
            Sheet sheet = baseSheet;
            assertFalse(sheet.isEmpty());
            assertEquals(1, sheet.getMinRowIndex());
            assertEquals(0, sheet.getMinColumnIndex());
            assertEquals(5, sheet.getMaxRowIndex());
            assertEquals(5, sheet.getMaxColumnIndex());
            assertEquals(6, sheet.getWidth());
            assertEquals(6, sheet.getDefinedWidth());
            assertEquals(6, sheet.getHeight());
            assertEquals(5, sheet.getDefinedHeight());
            assertEquals("Test cell 2", sheet.getCell(1, 2).text);
            assertEquals("#FF0000", sheet.getCellEntry(1, 2).getComputedFormat().get("background-color"));
            assertEquals("#FFFF00", sheet.getCellEntry(3, 1).getComputedFormat().get("background-color"));
            assertEquals("normal", sheet.getCellEntry(2, 2).getComputedFormat().get("font-weight"));
            assertNull(sheet.getColumn(1));
            assertNotNull(sheet.getColumn(2));
            assertNull(sheet.getRow(0));
            assertNotNull(sheet.getRow(1));
        }
        
        {
            Sheet sheet = new Sheet(baseSheet);
            assertFalse(sheet.isEmpty());
            assertEquals(1, sheet.getMinRowIndex());
            assertEquals(0, sheet.getMinColumnIndex());
            assertEquals(5, sheet.getMaxRowIndex());
            assertEquals(5, sheet.getMaxColumnIndex());
            assertEquals(6, sheet.getWidth());
            assertEquals(6, sheet.getDefinedWidth());
            assertEquals(6, sheet.getHeight());
            assertEquals(5, sheet.getDefinedHeight());
            assertEquals("Test cell 2", sheet.getCell(1, 2).text);
            assertEquals("#FF0000", sheet.getCellEntry(1, 2).getComputedFormat().get("background-color"));
            assertEquals("#FFFF00", sheet.getCellEntry(3, 1).getComputedFormat().get("background-color"));
            assertEquals("normal", sheet.getCellEntry(2, 2).getComputedFormat().get("font-weight"));
            assertNull(sheet.getColumn(1));
            assertNotNull(sheet.getColumn(2));
            assertNull(sheet.getRow(0));
            assertNotNull(sheet.getRow(1));
        }
    }


    @Test
    public void testSimpleMove() {
        Sheet sheet = new Sheet();
        sheet.write(-1,  -1, "Text");
        
        assertEquals(1, sheet.getWidth());
        assertEquals(1, sheet.getHeight());
        assertEquals(1, sheet.getDefinedWidth());
        assertEquals(1, sheet.getDefinedHeight());
        assertEquals("Text", sheet.getCell(-1, -1).text);
        assertNull(sheet.getCell(0, 0));
        assertNull(sheet.getCell(1, 2));
        
        int[] movedWith = sheet.moveToNonNegative();

        assertEquals(1, movedWith[0]);
        assertEquals(1, movedWith[1]);

        assertEquals(1, sheet.getWidth());
        assertEquals(1, sheet.getHeight());
        assertEquals(1, sheet.getDefinedWidth());
        assertEquals(1, sheet.getDefinedHeight());
        assertNull(sheet.getCell(-1, -1));
        assertEquals("Text", sheet.getCell(0, 0).text);
        assertNull(sheet.getCell(1, 2));
        
        sheet.move(1, 2);

        assertEquals(3, sheet.getWidth());
        assertEquals(2, sheet.getHeight());
        assertEquals(1, sheet.getDefinedWidth());
        assertEquals(1, sheet.getDefinedHeight());
        assertNull(sheet.getCell(-1, -1));
        assertNull(sheet.getCell(0, 0));
        assertEquals("Text", sheet.getCell(1, 2).text);
    }

    @Test
    public void testComplexMove() {
        Sheet sheet = createComplexSheet();

        assertFalse(sheet.isEmpty());
        assertEquals(1, sheet.getMinRowIndex());
        assertEquals(0, sheet.getMinColumnIndex());
        assertEquals(5, sheet.getMaxRowIndex());
        assertEquals(5, sheet.getMaxColumnIndex());
        assertEquals(6, sheet.getWidth());
        assertEquals(6, sheet.getDefinedWidth());
        assertEquals(6, sheet.getHeight());
        assertEquals(5, sheet.getDefinedHeight());
        assertEquals("Test cell 2", sheet.getCell(1, 2).text);
        assertEquals("#FF0000", sheet.getCellEntry(1, 2).getComputedFormat().get("background-color"));
        assertEquals("#FFFF00", sheet.getCellEntry(3, 1).getComputedFormat().get("background-color"));
        assertEquals("normal", sheet.getCellEntry(2, 2).getComputedFormat().get("font-weight"));
        assertNull(sheet.getColumn(1));
        assertNotNull(sheet.getColumn(2));
        assertNull(sheet.getRow(0));
        assertNotNull(sheet.getRow(1));
        
        sheet.move(-4, -7);
        
        assertFalse(sheet.isEmpty());
        assertEquals(-3, sheet.getMinRowIndex());
        assertEquals(-7, sheet.getMinColumnIndex());
        assertEquals(1, sheet.getMaxRowIndex());
        assertEquals(-2, sheet.getMaxColumnIndex());
        assertEquals(6, sheet.getWidth());
        assertEquals(6, sheet.getDefinedWidth());
        assertEquals(5, sheet.getHeight());
        assertEquals(5, sheet.getDefinedHeight());
        assertEquals("Test cell 2", sheet.getCell(-3, -5).text);
        assertEquals("#FF0000", sheet.getCellEntry(-3, -5).getComputedFormat().get("background-color"));
        assertEquals("#FFFF00", sheet.getCellEntry(-1, -6).getComputedFormat().get("background-color"));
        assertEquals("normal", sheet.getCellEntry(-1, -5).getComputedFormat().get("font-weight"));
        assertNull(sheet.getColumn(-6));
        assertNotNull(sheet.getColumn(-5));
        assertNull(sheet.getRow(-4));
        assertNotNull(sheet.getRow(-3));
        
        int[] move = sheet.moveToNonNegative();
        assertEquals(3, move[0]);
        assertEquals(7, move[1]);
        
        assertFalse(sheet.isEmpty());
        assertEquals(0, sheet.getMinRowIndex());
        assertEquals(0, sheet.getMinColumnIndex());
        assertEquals(4, sheet.getMaxRowIndex());
        assertEquals(5, sheet.getMaxColumnIndex());
        assertEquals(6, sheet.getWidth());
        assertEquals(6, sheet.getDefinedWidth());
        assertEquals(5, sheet.getHeight());
        assertEquals(5, sheet.getDefinedHeight());
        assertEquals("Test cell 2", sheet.getCell(0, 2).text);
        assertEquals("#FF0000", sheet.getCellEntry(0, 2).getComputedFormat().get("background-color"));
        assertEquals("#FFFF00", sheet.getCellEntry(2, 1).getComputedFormat().get("background-color"));
        assertEquals("normal", sheet.getCellEntry(1, 2).getComputedFormat().get("font-weight"));
        assertNull(sheet.getColumn(1));
        assertNotNull(sheet.getColumn(2));
        assertNull(sheet.getRow(-1));
        assertNotNull(sheet.getRow(0));
    }

    @Test
    public void testComplexCutAndInsert() {
        Sheet sheet = createComplexSheet();

        assertFalse(sheet.isEmpty());
        assertEquals(1, sheet.getMinRowIndex());
        assertEquals(0, sheet.getMinColumnIndex());
        assertEquals(5, sheet.getMaxRowIndex());
        assertEquals(5, sheet.getMaxColumnIndex());
        assertEquals(6, sheet.getWidth());
        assertEquals(6, sheet.getDefinedWidth());
        assertEquals(6, sheet.getHeight());
        assertEquals(5, sheet.getDefinedHeight());
        assertNull(sheet.getColumn(1));
        assertNotNull(sheet.getColumn(2));
        assertNull(sheet.getRow(0));
        assertNotNull(sheet.getRow(1));
        assertNull(sheet.getRow(3));
        assertNotNull(sheet.getRow(4));
        assertNotNull(sheet.getRow(5));
        {
            Sheet.Area area = sheet.areas.get(0);
            assertEquals(2, area.ranges.size());
            assertEquals(13, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            Sheet.Area area = sheet.areas.get(1);
            assertEquals(1, area.ranges.size());
            assertEquals(2, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            assertEquals(1, sheet.merges.size());
            Sheet.Range mergeRange = sheet.merges.get(0);
            assertEquals(1, mergeRange.rowIndex1);
            assertEquals(1, mergeRange.columnIndex1);
            assertEquals(2, mergeRange.rowIndex2);
            assertEquals(1, mergeRange.columnIndex2);
        }
        
        sheet.cutRow(2);

        assertFalse(sheet.isEmpty());
        assertEquals(1, sheet.getMinRowIndex());
        assertEquals(0, sheet.getMinColumnIndex());
        assertEquals(4, sheet.getMaxRowIndex());
        assertEquals(5, sheet.getMaxColumnIndex());
        assertEquals(6, sheet.getWidth());
        assertEquals(6, sheet.getDefinedWidth());
        assertEquals(5, sheet.getHeight());
        assertEquals(4, sheet.getDefinedHeight());
        assertNull(sheet.getColumn(1));
        assertNotNull(sheet.getColumn(2));
        assertNull(sheet.getRow(0));
        assertNotNull(sheet.getRow(1));
        assertNotNull(sheet.getRow(3));
        assertNotNull(sheet.getRow(4));
        assertNull(sheet.getRow(5));
        {
            Sheet.Area area = sheet.areas.get(0);
            assertEquals(2, area.ranges.size());
            assertEquals(13, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            Sheet.Area area = sheet.areas.get(1);
            assertEquals(1, area.ranges.size());
            assertEquals(2, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            assertEquals(1, sheet.merges.size());
            Sheet.Range mergeRange = sheet.merges.get(0);
            assertEquals(1, mergeRange.rowIndex1);
            assertEquals(1, mergeRange.columnIndex1);
            assertEquals(1, mergeRange.rowIndex2);
            assertEquals(1, mergeRange.columnIndex2);
        }

        sheet.cutColumn(1);

        assertFalse(sheet.isEmpty());
        assertEquals(1, sheet.getMinRowIndex());
        assertEquals(0, sheet.getMinColumnIndex());
        assertEquals(4, sheet.getMaxRowIndex());
        assertEquals(4, sheet.getMaxColumnIndex());
        assertEquals(5, sheet.getWidth());
        assertEquals(5, sheet.getDefinedWidth());
        assertEquals(5, sheet.getHeight());
        assertEquals(4, sheet.getDefinedHeight());
        assertNotNull(sheet.getColumn(1));
        assertNull(sheet.getColumn(2));
        assertNull(sheet.getRow(0));
        assertNotNull(sheet.getRow(1));
        assertNotNull(sheet.getRow(3));
        assertNotNull(sheet.getRow(4));
        assertNull(sheet.getRow(5));
        {
            Sheet.Area area = sheet.areas.get(0);
            assertEquals(2, area.ranges.size());
            assertEquals(11, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            Sheet.Area area = sheet.areas.get(1);
            assertEquals(1, area.ranges.size());
            assertEquals(1, IteratorUtils.toList(new Sheet.AreaPositionIterator(area)).size());
        }
        {
            assertEquals(0, sheet.merges.size());
        }
    }
    
    private Sheet createComplexSheet() {
        Sheet sheet = new Sheet();
        
        sheet.write(1, 1, "Test cell 1");
        sheet.write(1, 2, "Test cell 2", new Sheet.Format(new String[]{
            "font-weight", "bold",
            "background-color", "#FF0000"
        }));

        sheet.getColumn(2, new Sheet.Column()).format = new Sheet.Format(new String[]{
            "font-weight", "normal"
        });

        sheet.getRow(4, new Sheet.Row()).format = new Sheet.Format(new String[]{
            "font-style", "normal"
        });

        sheet.getRow(5, new Sheet.Row()).format = new Sheet.Format(new String[]{
            "font-weight", "normal"
        });
        
        {
            Sheet.Area area = new Sheet.Area(new int[]{
                1, 2, 1, 2,
                3, 0, 4, 5,
            }, new Sheet.Format(new String[]{
                "font-style", "italic",
                "background-color", "#FFFF00",
                "color", "#0000FF",
            }));
            assertFalse(area.contains(0, 1));
            assertTrue(area.contains(1, 2));
            assertFalse(area.contains(2, 2));
            assertTrue(area.contains(3, 2));
            assertTrue(area.contains(4, 5));
            assertFalse(area.contains(5, 6));
            sheet.areas.add(area);
        }
        
        sheet.areas.add(new Sheet.Area(3, 1, 3, 2, new Sheet.Format(new String[]{
            "color", "#CC9900",
            "text-decoration", "underline",
        })));
        
        sheet.merges.add(new Sheet.Range(1, 1, 2, 1));
        
        return sheet;
    }
    
}
