package org.example.dto;

import java.util.Comparator;

public class StringIdComporator implements Comparator<StringTable> {
    @Override
    public int compare(StringTable o1, StringTable o2) {
        return Integer.compare(o1.getId(), o2.getId());
    }
}
