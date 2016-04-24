package hu.webarticum.simple_spreadsheet_writer.util;

import static org.junit.Assert.*;

import org.junit.Test;


public class ColorUtilTest {

    @Test
    public void testParse() {
        {
            String colorValue = "#FF7733";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0xFF, parts[0]);
            assertEquals(0x77, parts[1]);
            assertEquals(0x33, parts[2]);
        }
        {
            String colorValue = "#452d34";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0x45, parts[0]);
            assertEquals(0x2D, parts[1]);
            assertEquals(0x34, parts[2]);
        }
        {
            String colorValue = "#347";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0x33, parts[0]);
            assertEquals(0x44, parts[1]);
            assertEquals(0x77, parts[2]);
        }
        {
            String colorValue = "#0000CC";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0, parts[0]);
            assertEquals(0, parts[1]);
            assertEquals(0xCC, parts[2]);
        }
        {
            String colorValue = "#1234";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0, parts[0]);
            assertEquals(0, parts[1]);
            assertEquals(0, parts[2]);
        }
        {
            String colorValue = "1234";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0, parts[0]);
            assertEquals(0, parts[1]);
            assertEquals(0, parts[2]);
        }
        {
            String colorValue = "#EFGHIJ";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0, parts[0]);
            assertEquals(0, parts[1]);
            assertEquals(0, parts[2]);
        }
        {
            String colorValue = "+!=+!(%";
            int[] parts = ColorUtil.parseColor(colorValue);
            assertEquals(0, parts[0]);
            assertEquals(0, parts[1]);
            assertEquals(0, parts[2]);
        }
    }

}
