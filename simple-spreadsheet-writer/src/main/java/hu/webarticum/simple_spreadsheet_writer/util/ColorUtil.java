package hu.webarticum.simple_spreadsheet_writer.util;


public class ColorUtil {

    static public int[] parseColor(String colorValue) {
        int red = 0;
        int green = 0;
        int blue = 0;
        if (colorValue.matches("#[a-fA-F0-9]{6}")) {
            red = Integer.parseInt(colorValue.substring(1, 3), 16);
            green = Integer.parseInt(colorValue.substring(3, 5), 16);
            blue = Integer.parseInt(colorValue.substring(5, 7), 16);
        } else if (colorValue.matches("#[a-fA-F0-9]{3}")) {
            char redChar = colorValue.charAt(1);
            char greenChar = colorValue.charAt(2);
            char blueChar = colorValue.charAt(3);
            red = Integer.parseInt("" + redChar + redChar, 16);
            green = Integer.parseInt("" + greenChar + greenChar, 16);
            blue = Integer.parseInt("" + blueChar + blueChar, 16);
        }
        return new int[]{red, green, blue};
    }
    
}
