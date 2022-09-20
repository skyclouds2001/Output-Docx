package org.sky.work;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class ToDoc {

    public static void exportDoc(ArrayList<Record> tableData, String[][] headers, String fileTitle, String fileOutputPath) throws IOException {

        XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(fileOutputPath);

        // 创建标题
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText(fileTitle);
        titleRun.setColor("000000");
        titleRun.setFontSize(20);
        titleRun.setBold(true);

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
                    cell = row.getCell(j);
                    paragraph = cell.getParagraphArray(0);
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    run = paragraph.createRun();
                    run.setBold(true);
                    run.setText(headers[i][j]);
                }
            }
        }

        // 插入数据
        for (int i = 0; i < tableData.size(); ++i) {
            XWPFTableRow row = table.getRow(i + 2);
            Record data = tableData.get(i);
            row.getCell(0).setText(data.type);
            row.getCell(1).setText(String.valueOf(data.index));
            row.getCell(2).setText(data.content);
            row.getCell(3).setText(data.desc);
            row.getCell(4).setText(data.imgURL);
            row.getCell(5).setText(data.advice);
            if (data.qType == 1) {
                row.getCell(6).setText("√");
            } else if (data.qType == 2) {
                row.getCell(7).setText("√");
            }
        }

        document.write(out);
        out.close();

    }

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
