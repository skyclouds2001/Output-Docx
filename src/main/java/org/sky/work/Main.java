package org.sky.work;

public class Main {
    public static void main(String[] args) {
        try {
            ToDoc.exportDoc(Data.records, Data.headers, "检查问题汇总清单", "D:\\程序\\outputPDF\\dist\\eg.docx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
