package hu.webarticum.simple_spreadsheet_writer;

import org.apache.commons.collections4.IteratorUtils;
import org.junit.Test;

import static org.junit.Assert.*;

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

        sheet.writeText(-1, -1, "Test text");

        assertTrue(sheet.hasNegative());
        assertEquals(1, sheet.getWidth());
        assertEquals(1, sheet.getHeight());

        sheet.removeCell(-1, -1);
        sheet.writeText(1, 1, "Test text");

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
    	Sheet sheet1 = new Sheet();
    	
    	// TODO: build sheet1

    	Sheet sheet2 = new Sheet(sheet1);
    	Sheet sheet3 = new Sheet(sheet2);
    	
    	Map<String, Sheet> sheetMap = new LinkedHashMap<String, Sheet>();
    	sheetMap.put("Sheet.1", sheet1);
    	sheetMap.put("Sheet.2", sheet2);
    	sheetMap.put("Sheet.3", sheet3);
    	
    	for (Map.Entry<String, Sheet> entry: sheetMap.entrySet()) {
    		String sheetName = entry.getKey();
    		Sheet sheet = entry.getValue();
    		try {
    		
    			// TODO: test sheet, assert with sheetName
    			
    		} catch (Throwable e) {
    			fail(sheetName + ": unexpected exception");
    			e.printStackTrace();
    		}
    	}
    }


    @Test
    public void testSimpleMove() {
    	Sheet sheet = new Sheet();
    	sheet.writeText(-1,  -1, "Text");
    	
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
    	// TODO
    }
    
}
