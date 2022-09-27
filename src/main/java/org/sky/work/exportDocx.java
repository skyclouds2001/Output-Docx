package org.sky.work;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.ServerSocket;
import java.net.Socket;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;
import java.util.Random;

public class exportDocx {
    private exportDocx() {
    }

    static void init() {
        exportDocx server = new exportDocx();
        server.listen();
    }

    private void listen() {
        try {
            try (ServerSocket serverSocket = new ServerSocket(8888)) {
                while (true) {
                    Socket client = serverSocket.accept();

                    new Thread(() -> {
                        try {
                            BufferedReader bd = new BufferedReader(new InputStreamReader(client.getInputStream()));
                            PrintWriter pw = new PrintWriter(new OutputStreamWriter(client.getOutputStream()));

                            try {
                                String[] requestLine = bd.readLine().split(" ");
                                @SuppressWarnings("unused")
                                String method = requestLine[0];
                                @SuppressWarnings("unused")
                                String path = requestLine[1];
                                @SuppressWarnings("unused")
                                String HTTP = requestLine[2];

                                if (Objects.equals(path, "/")) {
                                    String source = this.exportDoc("/dist/data.json");
                                    pw.println("HTTP/1.1 200 OK");
                                    pw.println("Content-type: text/plain;charset=utf-8");
                                    pw.println();
                                    pw.println(source);
                                }
                            } catch (Exception e) {
                                e.printStackTrace();
                                pw.println("HTTP/1.1 500 Internal Server Error");
                                pw.println("Content-type: text/plain;charset=utf-8");
                                pw.println();
                                pw.println("Fail");
                            }

                            pw.flush();
                            pw.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }).start();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 产生随机的文件名 - 随机数 + 时间戳
     *
     * @return 随机产生的文件名
     */
    private @NotNull String createFileName() {
        return new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + (new Random().nextInt(900) + 100);
    }

    /**
     * 判断图片文件的XWPFDocument中的类型
     *
     * @param imgFile 图片文件名|图片文件路径
     * @return 判断类型
     */
    private int getPictureFormat(@NotNull String imgFile) throws Exception {
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

    private @NotNull String getImageType(@NotNull String imgFile) throws Exception {
        String format;
        if (imgFile.endsWith(".emf")) format = "emf";
        else if (imgFile.endsWith(".wmf")) format = "wmf";
        else if (imgFile.endsWith(".pict")) format = "pict";
        else if (imgFile.endsWith(".jpg")) format = "jpg";
        else if (imgFile.endsWith(".jpeg")) format = "jpeg";
        else if (imgFile.endsWith(".png")) format = "png";
        else if (imgFile.endsWith(".dib")) format = "dib";
        else if (imgFile.endsWith(".gif")) format = "gif";
        else if (imgFile.endsWith(".tiff")) format = "tiff";
        else if (imgFile.endsWith(".eps")) format = "eps";
        else if (imgFile.endsWith(".bmp")) format = "bmp";
        else if (imgFile.endsWith(".wpg")) format = "wpg";
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
    private void mergeCellsVertically(@NotNull XWPFTable table, int col, int fromRow, int toRow) {
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
    private void mergeCellsHorizontal(@NotNull XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    private @NotNull String getFileName(@NotNull String path) {
        String[] strings = path.split("/");
        int length = strings.length;
        return strings[length - 1];
    }

    /**
     * 导出DOCX文档
     */
    private @NotNull String exportDoc(String path) throws Exception {
        // 初始化数据源
        String dataPath = System.getProperty("user.dir") + path;
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
        String filePath = System.getProperty("user.dir") + "\\dist\\" + createFileName() + ".docx";
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
        {
            String[][] headers = {
                    {"专业", "序号", "检查表中检查内容", "问题描述", "问题照片", "整改意见", "问题归属", "0"},
                    {"1", "1", "1", "1", "1", "1", "文件、资料类", "现场类"},
            };

            XWPFTableRow row_0 = table.getRow(0);
            for (int i = 0; i < 8 - 1; ++i) {
                row_0.addNewTableCell();
            }
            for (int i = 0; i < headers.length - 1; ++i) {
                table.createRow();
            }

            for (int i = 0; i < 2; ++i) {
                XWPFTableRow row = table.getRow(i);
                for (int j = 0; j < 8; ++j) {
                    if (Objects.equals(headers[i][j], "0")) {
                        mergeCellsHorizontal(table, i, j - 1, j);
                    } else if (Objects.equals(headers[i][j], "1")) {
                        mergeCellsVertically(table, j, i - 1, i);
                    } else {
                        XWPFTableCell cell = row.getCell(j);
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        XWPFParagraph paragraph = cell.getParagraphArray(0);
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setBold(true);
                        run.setText(headers[i][j]);
                    }
                }
            }
        }

        int base = 2;

        // 读取数据
        JSONObject source = JSON.parseObject(jsonStr);
        JSONArray allItemList = source.getJSONArray("all_checkup_subject_item_list");
        for (int i = 0; i < allItemList.size(); ++i) {
            JSONObject subject = allItemList.getJSONObject(i);

            // ***** 专业
            String subjectTitle = subject.getString("checkup_subject_name");

            int index = 0;

            @SuppressWarnings("unused")
            String subjectId = subject.getString("checkup_subject_id");

            JSONArray itemList = subject.getJSONArray("checkup_item_list");

            for (int j = 0; j < itemList.size(); ++j) {
                JSONObject item = itemList.getJSONObject(j);

                // ***** 检查表中的检查内容
                String content = item.getString("checkup_item_evaluation_content");

                JSONArray inspectionList = item.getJSONArray("checkup_item_inspection_point_content_list");

                // 多个评估方式，在表格内应呈现为多行
                for (int k = 0; k < inspectionList.size(); ++k) {
                    // 创建表格行
                    XWPFTableRow row = table.createRow();
                    // 设置检查表中的检查内容
                    row.getCell(2).setText(content);
                    row.getCell(1).setText(String.valueOf(index + 1));
                    ++index;

                    JSONObject inspection = inspectionList.getJSONObject(k);

                    @SuppressWarnings("unused")
                    String inspectionId = inspection.getString("checkup_subject_item_evaluation_origin_id");

                    JSONObject inspectionType = inspection.getJSONObject("evaluation_method_type_info");

                    if (inspectionType != null) {
                        // ***** 问题归属  1-书面材料 2-现场询问 3-实地考察
                        int inspectionTypeId = inspectionType.getIntValue("evaluation_method_type_id");

                        // 设置问题归属
                        if (inspectionTypeId == 1) {
                            row.getCell(6).setText("√");
                        } else {
                            row.getCell(7).setText("√");
                        }

                        @SuppressWarnings("unused")
                        String inspectionTypeName = inspectionType.getString("evaluation_method_type_name");
                    }

                    JSONArray evaluationList = inspection.getJSONArray("expert_evaluation_list");

                    // 如果列表中评估项数量大于1，则把所有评估项的所有问题描述/整改建议/问题照片 放在同一个单元格内，以换行符分开
                    for (int l = 0; l < evaluationList.size(); ++l) {
                        JSONObject evaluation = evaluationList.getJSONObject(l);

                        @SuppressWarnings("unused")
                        String evaluationId = evaluation.getString("expert_evaluation_id");

                        // ***** 问题描述
                        String evaluationDescription = evaluation.getString("checkup_evaluation_problem_description");
                        // ***** 整改建议
                        String evaluationImproveSuggest = evaluation.getString("checkup_evaluation_improve_suggest");

                        if (l == 0) {
                            row.getCell(3).getParagraphArray(0).createRun().setText(evaluationDescription);
                            row.getCell(5).getParagraphArray(0).createRun().setText(evaluationImproveSuggest);
                        } else {
                            row.getCell(3).addParagraph().createRun().setText(evaluationDescription);
                            row.getCell(5).addParagraph().createRun().setText(evaluationImproveSuggest);
                        }

                        JSONArray evaluationImageList = evaluation.getJSONArray("checkup_subject_item_evaluation_image_list");

                        for (int m = 0; m < evaluationImageList.size(); ++m) {
                            JSONObject evaluationImage = evaluationImageList.getJSONObject(m);

                            // ***** 问题照片
                            String evaluationImageURL = evaluationImage.getString("source");

                            try {
                                String imgURL = System.getProperty("user.dir") + "\\dist\\" + getFileName(evaluationImageURL);
                                File image = new File(imgURL);
                                BufferedImage img;
                                if (!image.exists()) {
                                    URL url = new URL(evaluationImageURL);
                                    img = ImageIO.read(url);
                                    String type = getImageType(evaluationImageURL);
                                    ImageIO.write(img, type, image);
                                } else {
                                    img = ImageIO.read(image);
                                }
                                int width = img.getWidth(), height = img.getHeight();
                                int cWidth = 100, cHeight = 100 * height / width;
                                if (l == 0 && m == 0) {
                                    row.getCell(4).getParagraphArray(0).createRun().addPicture(
                                            new FileInputStream(imgURL),
                                            getPictureFormat(evaluationImageURL),
                                            createFileName(),
                                            Units.toEMU(cWidth),
                                            Units.toEMU(cHeight)
                                    );
                                } else {
                                    row.getCell(4).addParagraph().createRun().addPicture(
                                            new FileInputStream(imgURL),
                                            getPictureFormat(evaluationImageURL),
                                            createFileName(),
                                            Units.toEMU(cWidth),
                                            Units.toEMU(cHeight)
                                    );
                                }
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                }
            }

            table.getRow(base).getCell(0).setText(subjectTitle);
            mergeCellsVertically(table, 0, base, base + index - 1);

            base += index;
        }

        document.write(out);
        out.flush();
        out.close();

        return filePath;
    }

}
