package org.sky.work;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.util.Arrays;

public class Main {
    public static void main(String[] args) {
        String[][] headers = {
                {"专业", "序号", "检查表中检查内容", "问题描述", "问题照片", "整改意见", "问题归属", "0"},
                {"1", "1", "1", "1", "1", "1", "文件、资料类", "现场类"},
        };
        try {
            JSONObject source = ReadJSON.readJSONfile("D:\\程序\\outputPDF\\src\\main\\resources\\data.json");
            String title = (String) source.get("title");
            Object[] experts = ((JSONArray) source.get("experts")).toArray();
            Object[] records = ((JSONArray) source.get("records")).toArray();
            String outputPath = "\\dist\\";

            ObjectMapper objectMapper = new ObjectMapper();

            String path = ToDoc.exportDoc(
                    title,
                    Arrays.stream(experts).toArray(String[]::new),
                    headers,
                    new Record[0], // todo
                    outputPath
            );
            System.out.println(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
