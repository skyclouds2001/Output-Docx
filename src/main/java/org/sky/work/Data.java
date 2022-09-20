package org.sky.work;

import java.util.ArrayList;

class Record {
    public String type;
    public int index;
    public String content;
    public String desc;
    public String imgURL;
    public String advice;
    public int qType;

    Record(String type, int index, String content, String desc, String imgURL, String advice, int qType) {
        this.type = type;
        this.index = index;
        this.content = content;
        this.desc = desc;
        this.imgURL = imgURL;
        this.advice = advice;
        this.qType = qType;
    }
}

public class Data {
    public static ArrayList<Record> records;

    static {
        records = new ArrayList<>();
        records.add(new Record("总图设计", 1, "企业", "神木", "/", "完善", 1));
        records.add(new Record("总图设计", 2, "企业", "神木", "/", "完善", 2));
    }
}
