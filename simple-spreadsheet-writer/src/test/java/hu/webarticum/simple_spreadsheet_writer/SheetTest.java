package hu.webarticum.simple_spreadsheet_writer;

import org.apache.commons.collections4.IteratorUtils;
import org.junit.Test;

import static org.junit.Assert.*;

import java.util.List;

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
    }

    @Test
    public void testNegativePositionSheet() {
        Sheet sheet = new Sheet();
        sheet.writeText(-3, -4, "Test text 1");
        sheet.writeText(1, 3, "Test text 2");
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
            sheet.writeText(1, 2, "Test text 1");
            sheet.writeText(2, 3, "Test text 2");
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
            assertTrue(sheet.iterator().hasNext());
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
        // TODO
    }

}
