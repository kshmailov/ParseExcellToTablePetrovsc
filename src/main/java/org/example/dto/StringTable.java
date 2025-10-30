package org.example.dto;

import lombok.AccessLevel;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;
@Getter
@Setter
public class StringTable {
    int id;
    String shm;
    String ts;
    String formula;
    String po;
    String slice;
    String sign;
    String kpr;
    @Setter(AccessLevel.NONE)
    List<String> uvList = new ArrayList<>();

    public void addUv(String uv){
        uvList.add(uv);
    }

    @Override
    public String toString() {
        return "StringTable{\n" +
                "id=" + id +
                "\nshm='" + shm + '\'' +
                "\nts='" + ts + '\'' +
                "\nformula='" + formula + '\'' +
                "\npo='" + po + '\'' +
                "\nslice='" + slice + '\'' +
                "\nsign='" + sign + '\'' +
                "\nkpr='" + kpr + '\'' +
                "\nuvList=" + uvList +
                '}';
    }
}
