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
    public static String[][] headers = {
            {"专业", "序号", "检查表中检查内容", "问题描述", "问题照片", "整改意见", "问题归属", "0"},
            {"1", "1", "1", "1", "1", "1", "文件、资料类", "现场类"},
    };

    public static String[] experts = {
            "专家A", "专家B"
    };

    static {
        records = new ArrayList<>();
        records.add(new Record("总图设计", 1, "企业", "神木", null, "完善", 1));
        records.add(new Record("总图设计", 2, "企业", "神木", "D:\\程序\\outputPDF\\dist\\repository-open-graph-template.png", "完善", 2));
    }
}
