package org.sky.work;

public class Main {
    public static void main(String[] args) {
        try {
            ToDoc.exportDoc("检查问题汇总清单", Data.experts, Data.headers, Data.records, "\\dist\\");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
