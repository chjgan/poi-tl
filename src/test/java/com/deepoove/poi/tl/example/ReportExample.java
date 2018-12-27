package com.deepoove.poi.tl.example;

import com.deepoove.poi.CreateChartServiceImpl;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.jfree.data.category.CategoryDataset;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
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

        List<ModelData> modelDatas = new ArrayList<ModelData>();
        for (String title:titles){
            ModelData modelData = new ModelData();
            modelData.setTable(new MiniTableRenderData(Arrays.asList(cities,price)));
            modelData.setPicture(pictureRenderData);
            modelData.setTitle(title);
            modelDatas.add(modelData);
        }
        DocxRenderData model = new DocxRenderData(new File("F:\\1\\model.docx"), modelDatas );
        reportData.setModel(model);

        XWPFTemplate template = XWPFTemplate.compile("F:\\1\\report.docx").render(reportData);

        FileOutputStream out = new FileOutputStream("报告.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }

}
