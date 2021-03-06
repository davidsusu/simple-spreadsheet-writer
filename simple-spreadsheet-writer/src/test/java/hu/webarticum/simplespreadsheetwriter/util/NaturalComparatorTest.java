package hu.webarticum.simplespreadsheetwriter.util;

import static org.junit.Assert.*;

import org.junit.Test;

import hu.webarticum.simplespreadsheetwriter.util.NaturalComparator;

public class NaturalComparatorTest {

    @Test
    public void test() {
        Integer integer1 = 3;
        Integer integer2 = 4;
        NaturalComparator comparator = new NaturalComparator();
        assertEquals(integer1.compareTo(integer2), comparator.compare(integer1, integer2));

        Object object1 = new Object();
        Object object2 = new Object();
        try {
            comparator.compare(object1, object2);
            fail();
        } catch (ClassCastException e) {
        }

        try {
            comparator.compare(integer1, object2);
            fail();
        } catch (ClassCastException e) {
        }
    }

}
