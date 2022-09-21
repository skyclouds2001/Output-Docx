package org.sky.work;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class ReadJSON {
    private ReadJSON() {
        throw new RuntimeException("Unable to create this class!");
    }

    static JSONObject readJSONfile(String filePath) throws IOException {
        FileReader fileReader = new FileReader(filePath);
        Reader reader = new InputStreamReader(new FileInputStream(filePath), StandardCharsets.UTF_8);
        int ch;
        StringBuffer sb = new StringBuffer();
        while ((ch = reader.read()) != -1) {
            sb.append((char) ch);
        }
        fileReader.close();
        reader.close();
        String jsonStr = sb.toString();
        return JSON.parseObject(jsonStr);
    }
}
