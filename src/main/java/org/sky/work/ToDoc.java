package org.sky.work;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.*;

public class ToDoc {

    /**
     * 将数据转为DOC文档表格的主要方法
     *
     * @param tableData      表格内容数据
     * @param tableHeaders   表格表头数据
     * @param fileTitle      文档标题
     * @param fileOutputPath 文件输出路径
     */
    public static @NotNull String exportDoc(@NotNull String fileTitle, @NotNull String @NotNull [] experts, @NotNull String[][] tableHeaders, @NotNull ArrayList<Record> tableData, @NotNull String fileOutputPath) throws IOException {

        XWPFDocument document = new XWPFDocument();
        String filePath = System.getProperty("user.dir") + fileOutputPath + createFileName() + ".docx";
        FileOutputStream out = new FileOutputStream(filePath);

        // 创建标题
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText(fileTitle);
        titleRun.setColor("000000");
        titleRun.setFontSize(20);
        titleRun.setBold(true);

        // 添加空行
        document.createParagraph();

        // 添加标注
        StringBuilder ss = new StringBuilder("专家组成员：");
        for (String s : experts) {
            ss.append(s);
        }
        XWPFParagraph expert = document.createParagraph();
        XWPFRun expertsRun = expert.createRun();
        expertsRun.setText(ss.toString());
        expertsRun.setBold(true);

        // 创建表格
        XWPFTable table = document.createTable();

        // 初始化表格
        {
            XWPFTableRow row;
            XWPFTableCell cell;
            XWPFParagraph paragraph;
            XWPFRun run;

            row = table.getRow(0);
            for (int i = 0; i < 8 - 1; ++i) {
                row.addNewTableCell();
            }
            for (int i = 0; i < tableData.size() - 1; ++i) {
                table.createRow();
            }
            table.createRow();
            table.createRow();

            for (int i = 0; i < 2; ++i) {
                row = table.getRow(i);
                for (int j = 0; j < 8; ++j) {
                    if (Objects.equals(tableHeaders[i][j], "0")) {
                        mergeCellsHorizontal(table, i, j - 1, j);
                    } else if (Objects.equals(tableHeaders[i][j], "1")) {
                        mergeCellsVertically(table, j, i - 1, i);
                    } else {
                        cell = row.getCell(j);
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        paragraph = cell.getParagraphArray(0);
                        paragraph.setAlignment(ParagraphAlignment.CENTER);
                        run = paragraph.createRun();
                        run.setBold(true);
                        run.setText(tableHeaders[i][j]);
                    }
                }
            }

        }

        // 插入数据
        for (int i = 0; i < tableData.size(); ++i) {
            XWPFTableRow row = table.getRow(i + 2);
            Record data = tableData.get(i);
            row.getCell(0).setText(data.type);
            if (i == 0 || !Objects.equals(tableData.get(i - 1).type, tableData.get(i).type)) {
                int pos = i;
                while (pos + 1 < tableData.size() && Objects.equals(tableData.get(pos).type, tableData.get(pos + 1).type))
                    ++pos;
                mergeCellsVertically(table, 0, i + 2, pos + 2);
            }
            row.getCell(1).setText(String.valueOf(data.index));
            row.getCell(2).setText(data.content);
            row.getCell(3).setText(data.desc);
            if (data.imgURL != null) {
                for (String s : data.imgURL) {
                    XWPFParagraph p = row.getCell(4).getParagraphArray(0);
                    XWPFRun run = p.createRun();
                    try {
                        run.addPicture(
                                new FileInputStream(s),
                                getPictureFormat(s),
                                "",
                                Units.toEMU(75),
                                Units.toEMU(50)
                        );
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            } else {
                row.getCell(4).setText("/");
            }
            row.getCell(5).setText(data.advice);
            if (data.qType == 1) {
                row.getCell(6).setText("√");
            } else if (data.qType == 2) {
                row.getCell(7).setText("√");
            }
        }

        document.write(out);
        out.flush();
        out.close();

        return filePath;

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

    /**
     * 演示方法
     */
    @SuppressWarnings("unused")
    public void test() throws IOException, InvalidFormatException {

        XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream("D:\\程序\\outputPDF\\dist\\eg.docx");


        /*
          标题
         */
        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun titleParagraphRun = titleParagraph.createRun();

        titleParagraphRun.setText("Java PoI");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);

        /*
         * 段落
         */
        //段落
        XWPFParagraph firstParagraph = document.createParagraph();
        XWPFRun run = firstParagraph.createRun();
        run.setText("Java POI 生成word文件。");
        run.setColor("696969");
        run.setFontSize(16);

        //设置段落背景颜色
        CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
        cTShd.setVal(STShd.CLEAR);
        cTShd.setFill("97FFFF");

        //换行
        XWPFParagraph paragraph1 = document.createParagraph();
        XWPFRun paragraphRun1 = paragraph1.createRun();
        paragraphRun1.setText("\r");

        /*
         * 表格
         */
        //基本信息表格
        XWPFTable infoTable = document.createTable();
        //去表格边框
        infoTable.getCTTbl().getTblPr().unsetTblBorders();

        //列宽自动分割
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));

        //表格第一行
        XWPFTableRow infoTableRowOne = infoTable.getRow(0);
        infoTableRowOne.getCell(0).setText("职位");
        infoTableRowOne.addNewTableCell().setText(": Java 开发工程师");

        //表格第二行
        XWPFTableRow infoTableRowTwo = infoTable.createRow();
        infoTableRowTwo.getCell(0).setText("姓名");
        infoTableRowTwo.getCell(1).setText(": seawater");

        //表格第三行
        XWPFTableRow infoTableRowThree = infoTable.createRow();
        infoTableRowThree.getCell(0).setText("生日");
        infoTableRowThree.getCell(1).setText(": xxx-xx-xx");

        //表格第四行
        XWPFTableRow infoTableRowFour = infoTable.createRow();
        infoTableRowFour.getCell(0).setText("性别");
        infoTableRowFour.getCell(1).setText(": 男");

        //表格第五行
        XWPFTableRow infoTableRowFive = infoTable.createRow();
        infoTableRowFive.getCell(0).setText("现居地");
        infoTableRowFive.getCell(1).setText(": xx");

        /*
         * 图片
         */
        XWPFParagraph pic = document.createParagraph();
        pic.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun picRun = pic.createRun();
        List<String> filePath = new ArrayList<>();
        filePath.add("D:\\程序\\outputPDF\\dist\\repository-open-graph-template.png");
        filePath.add("D:\\程序\\outputPDF\\dist\\repository-open-graph-template.png");
        filePath.add("D:\\程序\\outputPDF\\dist\\repository-open-graph-template.png");
        filePath.add("D:\\程序\\outputPDF\\dist\\repository-open-graph-template.png");

        for (String str : filePath) {
            picRun.setText(str);
            picRun.addPicture(
                    new FileInputStream(str), XWPFDocument.PICTURE_TYPE_JPEG,
                    str,
                    Units.toEMU(450),
                    Units.toEMU(300)
            );
        }

        /*
         * 页眉和页脚
         */
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

        //添加页眉
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "ctpHeader";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        //设置为右对齐
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

        //添加页脚
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "ctpFooter";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);


        document.write(out);
        out.close();

    }

}
