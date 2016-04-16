package hu.webarticum.simple_spreadsheet_writer;

import org.apache.commons.collections4.IteratorUtils;
import org.junit.Test;

import hu.webarticum.simple_spreadsheet_writer.util.MergeSortedIterator;

import static org.junit.Assert.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

public class MergeSortedIteratorTest {

    @Test
    public void testWithIntegers() {

        {
            List<Iterator<Integer>> iterators = new ArrayList<Iterator<Integer>>();
            MergeSortedIterator<Integer> sortedIterator = new MergeSortedIterator<Integer>(iterators);
            List<Integer> result = IteratorUtils.toList(sortedIterator);
            List<Integer> expected = new ArrayList<Integer>();
            assertEquals(expected, result);
        }

        {
            List<Iterator<Integer>> iterators = new ArrayList<Iterator<Integer>>();
            iterators.add(Arrays.asList(5, 10).iterator());
            iterators.add(Arrays.asList(3, 20).iterator());
            MergeSortedIterator<Integer> sortedIterator = new MergeSortedIterator<Integer>(iterators);
            List<Integer> result = IteratorUtils.toList(sortedIterator);
            List<Integer> expected = Arrays.asList(3, 5, 10, 20);
            assertEquals(expected, result);
        }
        
        {
            List<Iterator<Integer>> iterators = new ArrayList<Iterator<Integer>>();
            iterators.add(Arrays.asList(1, 2, 3).iterator());
            iterators.add(Arrays.asList(2, 4, 6).iterator());
            iterators.add(Arrays.asList(5, 9).iterator());
            MergeSortedIterator<Integer> sortedIterator = new MergeSortedIterator<Integer>(iterators);
            List<Integer> result = IteratorUtils.toList(sortedIterator);
            List<Integer> expected = Arrays.asList(1, 2, 3, 4, 5, 6, 9);
            assertEquals(expected, result);
        }

        {
            List<Iterator<Integer>> iterators = new ArrayList<Iterator<Integer>>();
            iterators.add(Arrays.asList(-1).iterator());
            iterators.add(Arrays.asList(-2, 0, 100).iterator());
            iterators.add(Arrays.asList(-100, 0, 100, 101).iterator());
            iterators.add(Arrays.asList(98).iterator());
            MergeSortedIterator<Integer> sortedIterator = new MergeSortedIterator<Integer>(iterators);
            List<Integer> result = IteratorUtils.toList(sortedIterator);
            List<Integer> expected = Arrays.asList(-100, -2, -1, 0, 98, 100, 101);
            assertEquals(expected, result);
        }

        {
            List<Iterator<Integer>> iterators = new ArrayList<Iterator<Integer>>();
            iterators.add(Arrays.asList(-1).iterator());
            iterators.add(Arrays.asList(-2, 0, 100).iterator());
            iterators.add(Arrays.asList(-100, 0, 100, 100, 101).iterator());
            iterators.add(Arrays.asList(98).iterator());
            MergeSortedIterator<Integer> sortedIterator = new MergeSortedIterator<Integer>(iterators);
            List<Integer> result = IteratorUtils.toList(sortedIterator);
            List<Integer> expected = Arrays.asList(-100, -2, -1, 0, 98, 100, 100, 101);
            assertEquals(expected, result);
        }
        
    }
    
    @Test
    public void testWithCustomComparator() {
        Comparator<String> comparator = new Comparator<String>() {

            @Override
            public int compare(String str1, String str2) {
                if (str1.isEmpty() || str2.isEmpty()) {
                    return -str1.compareTo(str2);
                }
                return -Character.compare(Character.toLowerCase(str1.charAt(0)), Character.toLowerCase(str2.charAt(0)));
            }
            
        };

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = new ArrayList<String>();
            assertEquals(expected, result);
        }
        
        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("abcd").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("abcd");
            assertEquals(expected, result);
        }

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("abcd").iterator());
            iterators.add(Arrays.asList("AXZ").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("abcd");
            assertEquals(expected, result);
        }

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("AXZ").iterator());
            iterators.add(Arrays.asList("abcd").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("AXZ");
            assertEquals(expected, result);
        }

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("cott", "azzz").iterator());
            iterators.add(Arrays.asList("foo", "burr").iterator());
            iterators.add(Arrays.asList("bizz", "BAR", "").iterator());
            iterators.add(Arrays.asList("fizz").iterator());
            iterators.add(Arrays.asList("serg", "boff", "ASTER").iterator());
            iterators.add(Arrays.asList("").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("serg", "foo", "cott", "burr", "BAR", "azzz", "");
            assertEquals(expected, result);
        }
        
    }

    @Test
    public void testWithCustomMerger() {
        Comparator<String> comparator = new Comparator<String>() {

            @Override
            public int compare(String str1, String str2) {
                return Character.compare(str1.charAt(0), str2.charAt(0));
            }
            
        };
        
        MergeSortedIterator.Merger<String> merger = new MergeSortedIterator.Merger<String>() {

            @Override
            public String merge(List<String> items) {
                int maxNumber = Integer.MIN_VALUE;
                String result = null;
                for (String item: items) {
                    int number = Integer.parseInt(item.substring(2));
                    if (number > maxNumber) {
                        result = item;
                    }
                }
                return result;
            }
            
        };

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator, merger);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = new ArrayList<String>();
            assertEquals(expected, result);
        }

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("a.1").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator, merger);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("a.1");
            assertEquals(expected, result);
        }

        {
            List<Iterator<String>> iterators = new ArrayList<Iterator<String>>();
            iterators.add(Arrays.asList("a.1", "b.3", "d.2").iterator());
            iterators.add(Arrays.asList("b.2", "c.1", "c.2").iterator());
            iterators.add(Arrays.asList("b.2", "c.3", "d.5").iterator());
            iterators.add(Arrays.asList("a.1", "e.2").iterator());
            MergeSortedIterator<String> sortedIterator = new MergeSortedIterator<String>(iterators, comparator, merger);
            List<String> result = IteratorUtils.toList(sortedIterator);
            List<String> expected = Arrays.asList("a.1", "b.2", "c.3", "c.2", "d.5", "e.2");
            assertEquals(expected, result);
        }

    }
    
}

