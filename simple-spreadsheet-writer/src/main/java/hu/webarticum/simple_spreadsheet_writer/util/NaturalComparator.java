package hu.webarticum.simple_spreadsheet_writer.util;

import java.util.Comparator;

public class NaturalComparator implements Comparator<Object> {

    @SuppressWarnings({"unchecked", "rawtypes"})
    @Override
    public int compare(Object object1, Object object2) {
        return ((Comparable)object1).compareTo(object2);
    }

}
