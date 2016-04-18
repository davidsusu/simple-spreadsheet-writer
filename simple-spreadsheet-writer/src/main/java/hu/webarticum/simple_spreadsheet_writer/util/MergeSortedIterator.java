package hu.webarticum.simple_spreadsheet_writer.util;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.commons.collections4.iterators.PeekingIterator;


public class MergeSortedIterator<E> implements Iterator<E> {

    protected final List<PeekingIterator<E>> peekingIterators = new ArrayList<PeekingIterator<E>>();
    
    protected final Comparator<? super E> comparator;
    
    protected final Merger<E> merger;

    @SafeVarargs
    public MergeSortedIterator(Iterator<E>... iterators) {
        this(Arrays.asList(iterators), new NaturalComparator(), new DefaultMerger<E>());
    }

    @SuppressWarnings("unchecked")
    public MergeSortedIterator(Iterator<E> iterator1, Iterator<E> iterator2, Comparator<? super E> comparator) {
        this(Arrays.asList(iterator1, iterator2), comparator, new DefaultMerger<E>());
    }

    @SuppressWarnings("unchecked")
    public MergeSortedIterator(Iterator<E> iterator1, Iterator<E> iterator2, Comparator<? super E> comparator, Merger<E> merger) {
        this(Arrays.asList(iterator1, iterator2), comparator, merger);
    }

    public MergeSortedIterator(Collection<Iterator<E>> iterators) {
        this(iterators, new NaturalComparator(), new DefaultMerger<E>());
    }

    public MergeSortedIterator(Collection<Iterator<E>> iterators, Comparator<? super E> comparator) {
        this(iterators, comparator, new DefaultMerger<E>());
    }
    
    public MergeSortedIterator(Collection<Iterator<E>> iterators, Comparator<? super E> comparator, Merger<E> merger) {
        for (Iterator<E> iterator: iterators) {
            this.peekingIterators.add(new PeekingIterator<E>(iterator));
        }
        this.comparator = comparator;
        this.merger = merger;
    }

    @Override
    public boolean hasNext() {
        for (PeekingIterator<E> iterator: peekingIterators) {
            if (iterator.hasNext()) {
                return true;
            }
        }
        return false;
    }

    @Override
    public E next() {
        List<E> items = new ArrayList<E>();
        for (PeekingIterator<E> iterator: peekingIterators) {
            if (iterator.hasNext()) {
                E item = iterator.peek();
                if (items.isEmpty()) {
                    items.add(item);
                } else {
                    E referenceItem = items.get(0);
                    int cmp = comparator.compare(item, referenceItem);
                    if (cmp == 0) {
                        items.add(item);
                    } else if (cmp < 0) {
                        items.clear();
                        items.add(item);
                    }
                }
            }
        }
        E referenceItem = items.get(0);
        if (items.isEmpty()) {
            throw new NoSuchElementException();
        }
        for (PeekingIterator<E> iterator: peekingIterators) {
            if (iterator.hasNext()) {
                E item = iterator.peek();
                int cmp = comparator.compare(item, referenceItem);
                if (cmp == 0) {
                    iterator.next();
                }
            }
        }
        return merger.merge(items);
    }

    @Override
    public void remove() {
        throw new UnsupportedOperationException();
    }
    
    public interface Merger<E> {
        
        public E merge(List<E> items);
        
    }
    
    static public class DefaultMerger<E> implements Merger<E> {

        @Override
        public E merge(List<E> items) {
            if (items.isEmpty()) {
                return null;
            } else {
                return items.get(0);
            }
        }
        
    }
    
}
