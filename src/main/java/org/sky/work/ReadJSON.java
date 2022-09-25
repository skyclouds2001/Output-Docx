package org.sky.work;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Random;

public class ReadJSON {
    private ReadJSON() {
        throw new RuntimeException("Unable to create this class!");
    }

    static void run() {
        try {
            String path = exportDoc();
            System.out.println(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 产生随机的文件名 - 随机数 + 时间戳
     *
     * @return 随机产生的文件名
     */
    private static @NotNull String createFileName() {
        return new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + (new Random().nextInt(900) + 100);
    }

    /**
     * 判断图片文件的XWPFDocument中的类型
     *
     * @param imgFile 图片文件名|图片文件路径
     * @return 判断类型
     */
    private static int getPictureFormat(@NotNull String imgFile) throws Exception {
        int format;
        if (imgFile.endsWith(".emf")) format = XWPFDocument.PICTURE_TYPE_EMF;
        else if (imgFile.endsWith(".wmf")) format = XWPFDocument.PICTURE_TYPE_WMF;
        else if (imgFile.endsWith(".pict")) format = XWPFDocument.PICTURE_TYPE_PICT;
        else if (imgFile.endsWith(".jpeg") || imgFile.endsWith(".jpg")) format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if (imgFile.endsWith(".png")) format = XWPFDocument.PICTURE_TYPE_PNG;
        else if (imgFile.endsWith(".dib")) format = XWPFDocument.PICTURE_TYPE_DIB;
        else if (imgFile.endsWith(".gif")) format = XWPFDocument.PICTURE_TYPE_GIF;
        else if (imgFile.endsWith(".tiff")) format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if (imgFile.endsWith(".eps")) format = XWPFDocument.PICTURE_TYPE_EPS;
        else if (imgFile.endsWith(".bmp")) format = XWPFDocument.PICTURE_TYPE_BMP;
        else if (imgFile.endsWith(".wpg")) format = XWPFDocument.PICTURE_TYPE_WPG;
        else {
            throw new Exception("不支持的图片格式: " + imgFile + ". 仅支持 emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg 格式的图片");
        }
        return format;
    }

    /**
     * 垂直方向合并单元格
     *
     * @param table   表格对象本身
     * @param col     列数
     * @param fromRow 开始行数
     * @param toRow   结束行数
     */
    private static void mergeCellsVertically(@NotNull XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * 水平方向合并单元格
     *
     * @param table    表格对象本身
     * @param row      行数
     * @param fromCell 开始列数
     * @param toCell   结束列数
     */
    private static void mergeCellsHorizontal(@NotNull XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    static @NotNull String exportDoc() throws IOException {
        // 初始化数据源
        String dataPath = System.getProperty("user.dir") + "/dist/data.json";
        Reader reader = new InputStreamReader(new FileInputStream(dataPath), StandardCharsets.UTF_8);
        int ch;
        StringBuilder sb = new StringBuilder();
        while ((ch = reader.read()) != -1) {
            sb.append((char) ch);
        }
        reader.close();
        String jsonStr = sb.toString();

        // 初始化文档
        XWPFDocument document = new XWPFDocument();
        String filePath = System.getProperty("user.dir") + "/dist/" + createFileName() + ".docx";
        FileOutputStream out = new FileOutputStream(filePath);

        // 创建标题
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("检查问题汇总清单");
        titleRun.setColor("000000");
        titleRun.setFontSize(20);
        titleRun.setBold(true);

        // 添加空行
        document.createParagraph();

        // 创建表格
        XWPFTable table = document.createTable();

        // 初始化表头
        XWPFTableRow row_0 = table.getRow(0);
        XWPFTableCell cell_0_0 = row_0.getCell(0);
        cell_0_0.setText("专业");
        cell_0_0.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_0 = cell_0_0.getParagraphArray(0);
        p_0_0.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_1 = row_0.createCell();
        cell_0_1.setText("序号");
        cell_0_1.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_1 = cell_0_1.getParagraphArray(0);
        p_0_1.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_2 = row_0.createCell();
        cell_0_2.setText("检查表中检查内容");
        cell_0_2.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_2 = cell_0_2.getParagraphArray(0);
        p_0_2.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_3 = row_0.createCell();
        cell_0_3.setText("问题描述");
        cell_0_3.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_3 = cell_0_3.getParagraphArray(0);
        p_0_3.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_4 = row_0.createCell();
        cell_0_4.setText("问题照片");
        cell_0_4.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_4 = cell_0_4.getParagraphArray(0);
        p_0_4.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_5 = row_0.createCell();
        cell_0_5.setText("整改建议");
        cell_0_5.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_5 = cell_0_5.getParagraphArray(0);
        p_0_5.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_0_6 = row_0.createCell();
        cell_0_6.setText("问题归属");
        cell_0_6.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_0_6 = cell_0_6.getParagraphArray(0);
        p_0_6.setAlignment(ParagraphAlignment.CENTER);
        row_0.createCell();
        mergeCellsHorizontal(table, 0, 6, 7);
        XWPFTableRow row_1 = table.createRow();
        XWPFTableCell cell_1_6 = row_1.getCell(6);
        cell_1_6.setText("文件、 资料类");
        cell_1_6.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_1_6 = cell_1_6.getParagraphArray(0);
        p_1_6.setAlignment(ParagraphAlignment.CENTER);
        XWPFTableCell cell_1_7 = row_1.getCell(7);
        cell_1_7.setText("现场类");
        cell_1_7.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        XWPFParagraph p_1_7 = cell_1_7.getParagraphArray(0);
        p_1_7.setAlignment(ParagraphAlignment.CENTER);
        mergeCellsVertically(table, 0, 0, 1);
        mergeCellsVertically(table, 1, 0, 1);
        mergeCellsVertically(table, 2, 0, 1);
        mergeCellsVertically(table, 3, 0, 1);
        mergeCellsVertically(table, 4, 0, 1);
        mergeCellsVertically(table, 5, 0, 1);

        // 读取数据
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

                // 多个评估方式，在表格内应呈现为多行
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

                    // 如果列表中评估项数量大于1，则把所有评估项的所有问题描述/整改建议/问题照片 放在同一个单元格内，以换行符分开
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

        document.write(out);
        out.flush();
        out.close();

        return filePath;
    }

}
