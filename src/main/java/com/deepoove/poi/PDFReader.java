package com.deepoove.poi;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfReaderContentParser;
import com.itextpdf.text.pdf.parser.SimpleTextExtractionStrategy;
import com.itextpdf.text.pdf.parser.TextExtractionStrategy;

import java.io.IOException;
import java.util.Arrays;

public class PDFReader {

    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws Exception {
//        System.out.print(getPdfFileText("F:\\1\\pdf1.pdf"));
        String[] titles = {"会话数量","会话时长","指令数量","数据库会话数量","作业执行数量","任务执行数量"};
        PdfReader reader = new PdfReader("F:\\1\\pdf1.pdf");
        PdfReaderContentParser parser = new PdfReaderContentParser(reader);
        Arrays.stream(titles).forEach(t->getPage(t,reader,parser));
    }

    public static String getPdfFileText(String fileName) throws IOException {
        PdfReader reader = new PdfReader(fileName);
        PdfReaderContentParser parser = new PdfReaderContentParser(reader);
        StringBuffer buff = new StringBuffer();
        TextExtractionStrategy strategy;
        for (int i = 1; i <= reader.getNumberOfPages(); i++) {
            strategy = parser.processContent(i,
                    new SimpleTextExtractionStrategy());
            buff.append(strategy.getResultantText());
        }
        return buff.toString();
    }

    public static int getPage(String title,PdfReader reader,PdfReaderContentParser parser) {
        try {
            for (int i = 1; i <= reader.getNumberOfPages(); i++) {
                String text = parser.processContent(i, new SimpleTextExtractionStrategy()).getResultantText();
                if (text.contains(title) && text.contains("详细数据")){
                    System.out.println(title+"第"+i+"页");
                    return i;
                }
            }
            return 0;
        }catch (Exception e){
            return 0;
        }

    }

}