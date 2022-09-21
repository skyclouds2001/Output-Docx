package org.sky.work;

public class Main {
    public static void main(String[] args) {
        try {
            ToDoc.exportDoc(Data.records, Data.headers, "检查问题汇总清单", "\\dist\\");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
