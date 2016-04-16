package hu.webarticum.simple_spreadsheet_writer.util;

import java.util.Comparator;

public class NaturalComparator<T> implements Comparator<T> {

    @SuppressWarnings({"unchecked", "rawtypes"})
    @Override
    public int compare(T object1, T object2) {
        return ((Comparable)object1).compareTo(object2);
    }

}
