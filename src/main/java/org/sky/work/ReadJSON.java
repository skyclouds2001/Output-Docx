package org.sky.work;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;

public class ReadJSON {
    private ReadJSON() {
        throw new RuntimeException("Unable to create this class!");
    }

    static void run() {
        try {
            readJSONfile();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static String[][] headers = {
            {"专业", "序号", "检查表中检查内容", "问题描述", "问题照片", "整改意见", "问题归属", "0"},
            {"1", "1", "1", "1", "1", "1", "文件、资料类", "现场类"},
    };

    static void readJSONfile() throws IOException {
        String dataPath = System.getProperty("user.dir") + "/dist/data.json";
        Reader reader = new InputStreamReader(new FileInputStream(dataPath), StandardCharsets.UTF_8);
        int ch;
        StringBuilder sb = new StringBuilder();
        while ((ch = reader.read()) != -1) {
            sb.append((char) ch);
        }
        reader.close();
        String jsonStr = sb.toString();

        JSONObject source = JSON.parseObject(jsonStr);

        JSONArray allItemList = source.getJSONArray("all_checkup_subject_item_list");

        for (int i = 0; i < allItemList.size(); ++i) {
            JSONObject subject = allItemList.getJSONObject(i);

            // todo 专业
            String subjectTitle = subject.getString("checkup_subject_name");

            String subjectId = subject.getString("checkup_subject_id");

            JSONArray itemList = subject.getJSONArray("checkup_item_list");

            for (int j = 0; j < itemList.size(); ++j) {
                JSONObject item = itemList.getJSONObject(j);

                // todo 检查表中的检查内容
                String content = item.getString("checkup_item_evaluation_content");

                JSONArray inspectionList = item.getJSONArray("checkup_item_inspection_point_content_list");

                for (int k = 0; k < inspectionList.size(); ++k) {
                    JSONObject inspection = inspectionList.getJSONObject(k);

                    String inspectionId = inspection.getString("checkup_subject_item_evaluation_origin_id");

                    JSONObject inspectionType = inspection.getJSONObject("evaluation_method_type_info");

                    if (inspectionType != null) {
                        // todo 问题归属  1-书面材料 2-现场询问 3-实地考察
                        String inspectionTypeId = inspectionType.getString("evaluation_method_type_id");

                        String inspectionTypeName = inspectionType.getString("evaluation_method_type_name");
                    }

                    JSONArray evaluationList = inspection.getJSONArray("expert_evaluation_list");

                    for (int l = 0; l < evaluationList.size(); ++l) {
                        JSONObject evaluation = evaluationList.getJSONObject(l);

                        String evaluationId = evaluation.getString("expert_evaluation_id");

                        // todo 问题描述
                        String evaluationDescription = evaluation.getString("");
                        // todo 整改建议
                        String evaluationImproveSuggest = evaluation.getString("checkup_evaluation_improve_suggest");

                        JSONArray evaluationImageList = evaluation.getJSONArray("checkup_subject_item_evaluation_image_list");

                        for (int m = 0; m < evaluationImageList.size(); ++m) {
                            JSONObject evaluationImage = evaluationImageList.getJSONObject(m);

                            // todo 问题照片
                            String evaluationImageURL = evaluationImage.getString("source");
                        }
                    }
                }
            }
        }
    }
}
