package com.deepoove.poi.tl.example;

import com.deepoove.poi.AsposeToPDF;
import com.deepoove.poi.CreateChartServiceImpl;
import com.deepoove.poi.PDFReader;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfReaderContentParser;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.jfree.data.category.CategoryDataset;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

public class ReportExample {

    @Test
    public void myDetailTest() throws Exception{

        double[][] data1 = new double[][] { { 672, 766, 223, 540, 126 } };
        String[] rowKeys = { "苹果" };
        String[] columnKeys = { "北京", "上海", "广州", "成都", "深圳" };
        CategoryDataset dataset = CreateChartServiceImpl.getBarData(data1, rowKeys, columnKeys);
        final PictureRenderData pictureRenderData = new PictureRenderData(510, 412, ".jpg", BytePictureUtils.getBufferByteArray(CreateChartServiceImpl.getBarChart(dataset, "x坐标", "y坐标", "柱状图", "bar.png")));

        String[] row0 = new String[]{"产地"};
        row0 = ArrayUtils.addAll(row0,columnKeys);
        final RowRenderData cities = RowRenderData.build(row0);

        String[] row1 = {"价格"};
        for (double num:data1[0]){
            row1 = ArrayUtils.addAll(row1,String.valueOf(num));
        }
        final RowRenderData price = RowRenderData.build(row1);


        XWPFTemplate template = XWPFTemplate.compile("F:\\1\\detail.docx").render(new HashMap<String, Object>(){{
            put("title", "会话数量");
            put("picture",pictureRenderData);
            put("table",new MiniTableRenderData(Arrays.asList(cities,price)));
        }});

        FileOutputStream out = new FileOutputStream("detail.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }

    @Test
    public void myReportTest() throws Exception{
        String[] titles = {"会话数量","会话时长","指令数量","数据库会话数量","作业执行数量","任务执行数量"};
        ReportData reportData = new ReportData();
        reportData.setCreateTime("2018年12月27日");
        reportData.setDataTime("2018年12月20日");
        reportData.setNames("张小强、高小峰、王小明");
        StringBuilder stringBuilder = new StringBuilder();
        for (String key:titles){
            stringBuilder.append(key).append("、");
        }
        reportData.setTitles(stringBuilder.deleteCharAt(stringBuilder.length()-1).toString());

        double[][] data1 = new double[][] { { 672, 766, 223, 540, 126 } };
        String[] rowKeys = { "苹果" };
        String[] columnKeys = { "北京", "上海", "广州", "成都", "深圳" };
        CategoryDataset dataset = CreateChartServiceImpl.getBarData(data1, rowKeys, columnKeys);
        final PictureRenderData pictureRenderData = new PictureRenderData(510, 412, ".jpg", BytePictureUtils.getBufferByteArray(CreateChartServiceImpl.getBarChart(dataset, "x坐标", "y坐标", "柱状图", "bar.png")));

        String[] row0 = new String[]{"产地"};
        row0 = ArrayUtils.addAll(row0,columnKeys);
        final RowRenderData cities = RowRenderData.build(row0);

        String[] row1 = {"价格"};
        for (double num:data1[0]){
            row1 = ArrayUtils.addAll(row1,String.valueOf(num));
        }
        final RowRenderData price = RowRenderData.build(row1);

//        List<ModelData> modelDatas = new ArrayList<ModelData>();
//        for (String title:titles){
//            ModelData modelData = new ModelData();
//            modelData.setTable(new MiniTableRenderData(Arrays.asList(cities,price)));
//            modelData.setPicture(pictureRenderData);
//            modelDatas.add(modelData);
//        }

        reportData.setPicture1(pictureRenderData);
        reportData.setPicture2(pictureRenderData);
        reportData.setPicture3(pictureRenderData);
        reportData.setPicture4(pictureRenderData);
        reportData.setPicture5(pictureRenderData);
        reportData.setPicture6(pictureRenderData);

        MiniTableRenderData table = new MiniTableRenderData(Arrays.asList(cities, price));
        reportData.setTable1(table);
        reportData.setTable2(table);
        reportData.setTable3(table);
        reportData.setTable4(table);
        reportData.setTable5(table);
        reportData.setTable6(table);
//        DocxRenderData model = new DocxRenderData(new File("F:\\1\\model.docx"), modelDatas );
//        reportData.setModel(model);

        XWPFTemplate template = XWPFTemplate.compile("F:\\1\\report.docx").render(reportData);
        List<XWPFParagraph> paragraphs = template.getXWPFDocument().getParagraphs();
        int num=0;
        for (XWPFParagraph paragraph:paragraphs){
            CTP ctp = paragraph.getCTP();
            System.out.println(paragraph.getText());
            for (CTR ctr:ctp.getRList()){
//                ctr.getTList().forEach(t-> System.out.println(t.getStringValue()));
                if (ctr.toString().contains("lastRenderedPageBreak")){
                    System.out.println("有效页："+ctr.toString());
                    num++;
                }
//                System.out.println("有效页："+ctr.toString());
            }
        }
        System.out.println("总页数："+num);
//        List<XWPFParagraph> paragraphList = paragraphs.stream().filter(p -> p.getCTP().getHyperlinkList().size() > 0).collect(Collectors.toList());
//        List<XWPFParagraph> collect = paragraphs.stream().filter(p -> p.getCTP().getRList().stream().anyMatch(ctr -> ctr.getLastRenderedPageBreakList().size() > 0)).collect(Collectors.toList());
//        collect.forEach(p->p.getCTP().getRList().forEach(ctr -> ctr.getTList().forEach(t-> System.out.println(t.getStringValue()))));
//        paragraphList.forEach(p->p.getCTP().getHyperlinkArray(0).getRArray(p.getCTP().getHyperlinkArray(0).getRList().size()-2).getTArray(0).setStringValue("1"));

        FileOutputStream out = new FileOutputStream("F:\\1\\报告.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }

    @Test
    public void myReportFinal() throws Exception{
        String[] titles = {"会话数量","会话时长","指令数量","数据库会话数量","作业执行数量","任务执行数量"};
        ReportData reportData = new ReportData();
        reportData.setCreateTime("2018年12月27日");
        reportData.setDataTime("2018年12月20日");
        reportData.setNames("张小强、高小峰、王小明");
        StringBuilder stringBuilder = new StringBuilder();
        for (String key:titles){
            stringBuilder.append(key).append("、");
        }
        reportData.setTitles(stringBuilder.deleteCharAt(stringBuilder.length()-1).toString());

        double[][] data1 = new double[][] { { 672, 766, 223, 540, 126 } };
        String[] rowKeys = { "苹果" };
        String[] columnKeys = { "北京", "上海", "广州", "成都", "深圳" };
        CategoryDataset dataset = CreateChartServiceImpl.getBarData(data1, rowKeys, columnKeys);
        final PictureRenderData pictureRenderData = new PictureRenderData(510, 412, ".jpg", BytePictureUtils.getBufferByteArray(CreateChartServiceImpl.getBarChart(dataset, "x坐标", "y坐标", "柱状图", "bar.png")));

        String[] row0 = new String[]{"产地"};
        row0 = ArrayUtils.addAll(row0,columnKeys);
        final RowRenderData cities = RowRenderData.build(row0);

        String[] row1 = {"价格"};
        for (double num:data1[0]){
            row1 = ArrayUtils.addAll(row1,String.valueOf(num));
        }
        final RowRenderData price = RowRenderData.build(row1);


        reportData.setPicture1(pictureRenderData);
        reportData.setPicture2(pictureRenderData);
        reportData.setPicture3(pictureRenderData);
        reportData.setPicture4(pictureRenderData);
        reportData.setPicture5(pictureRenderData);
        reportData.setPicture6(pictureRenderData);

        MiniTableRenderData table = new MiniTableRenderData(Arrays.asList(cities, price));
        reportData.setTable1(table);
        reportData.setTable2(table);
        reportData.setTable3(table);
        reportData.setTable4(table);
        reportData.setTable5(table);
        reportData.setTable6(table);


        File tmp = File.createTempFile("tmp", ".docx", createDir("F:\\2"));
        File tmpPDF = File.createTempFile("tmp", ".pdf", createDir("F:\\2"));

        XWPFTemplate template = XWPFTemplate.compile("F:\\1\\report.docx").render(reportData);
        FileOutputStream out = new FileOutputStream(tmp);
        template.write(out);
        out.flush();
        out.close();

        AsposeToPDF.doc2pdf(tmp,tmpPDF);
        PdfReader reader = new PdfReader(tmpPDF.getPath());
        PdfReaderContentParser parser = new PdfReaderContentParser(reader);
        for (String title:titles){
            int page = PDFReader.getPage(title, reader, parser);
            template.getXWPFDocument().getParagraphs().stream().filter(p -> p.getCTP().getHyperlinkList().size() > 0 && p.getText().contains(title)).forEach(
                    p->p.getCTP().getHyperlinkArray(0).getRArray(p.getCTP().getHyperlinkArray(0).getRList().size()-2).getTArray(0).setStringValue(String.valueOf(page)));
        }
        File fdoc = File.createTempFile("tmp", ".docx", createDir("F:\\2"));
        FileOutputStream fout = new FileOutputStream(fdoc);
        template.write(fout);
        fout.flush();
        fout.close();
        template.close();
        tmp.deleteOnExit();
        tmpPDF.deleteOnExit();

    }

    /**
     * 创建目录
     * @param destDirName 目标目录名
     */

    public static File createDir(String destDirName) throws IOException {
        File dir = new File(destDirName);
        File parent = dir.getParentFile();
        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }
        if (!dir.exists()) {
            dir.mkdir();
        }
        return dir;
    }
}
