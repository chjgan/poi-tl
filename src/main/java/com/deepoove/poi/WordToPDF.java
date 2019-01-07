package com.deepoove.poi;

import com.lowagie.text.pdf.PdfReader;
import com.lowagie.text.pdf.SimpleBookmark;
import org.apache.commons.collections.MapUtils;
import org.apache.pdfbox.io.RandomAccessBuffer;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineNode;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.converter.core.utils.StringUtils;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


public class WordToPDF {


    /**
     * 将word文档， 转换成pdf, 中间替换掉变量
     * @param source 源为word文档， 必须为docx文档
     * @param target 目标输出
     * @param params 需要替换的变量
     * @throws Exception
     */
    public static void wordConverterToPdf(InputStream source,
                                          OutputStream target, Map<String, String> params) throws Exception {
        wordConverterToPdf(source, target, null, params);
    }

    /**
     * 将word文档， 转换成pdf, 中间替换掉变量
     * @param source 源为word文档， 必须为docx文档
     * @param target 目标输出
     * @param params 需要替换的变量
     * @param options PdfOptions.create().fontEncoding( "windows-1250" ) 或者其他
     * @throws Exception
     */
    public static void wordConverterToPdf(InputStream source, OutputStream target,
                                          PdfOptions options,
                                          Map<String, String> params) throws Exception {
        XWPFDocument doc = new XWPFDocument(source);
        paragraphReplace(doc.getParagraphs(), params);
        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    paragraphReplace(cell.getParagraphs(), params);
                }
            }
        }
        PdfConverter.getInstance().convert(doc, target, options);
    }

    /** 替换段落中内容 */
    private static void paragraphReplace(List<XWPFParagraph> paragraphs, Map<String, String> params) {
        if (MapUtils.isNotEmpty(params)) {
            for (XWPFParagraph p : paragraphs){
                for (XWPFRun r : p.getRuns()){
                    String content = r.getText(r.getTextPosition());
                    if(StringUtils.isNotEmpty(content) && params.containsKey(content)) {
                        r.setText(params.get(content), 0);
                    }
                }
            }
        }
    }

    /**
     * 将word文档， 转换成pdf
     * 宋体：STSong-Light
     *
     * @param fontParam1 可以字体的路径，也可以是itextasian-1.5.2.jar提供的字体，比如宋体"STSong-Light"
     * @param fontParam2 和fontParam2对应，fontParam1为路径时，fontParam2=BaseFont.IDENTITY_H，为itextasian-1.5.2.jar提供的字体时，fontParam2="UniGB-UCS2-H"
     * @param tmp        源为word文档， 必须为docx文档
     * @param target     目标输出
     * @throws Exception
     */
    public static void wordConverterToPdf(String tmp, String target, final String fontParam1, final String fontParam2) {
        InputStream sourceStream = null;
        OutputStream targetStream = null;
        XWPFDocument doc = null;
        try {
            sourceStream = new FileInputStream(tmp);
            targetStream = new FileOutputStream(target);
            doc = new XWPFDocument(sourceStream);
            PdfOptions options = PdfOptions.create();
            /*//中文字体处理
            options.fontProvider(new IFontProvider() {
                public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                    try {
                        BaseFont bfChinese = BaseFont.createFont(fontParam1, fontParam2, BaseFont.NOT_EMBEDDED);
                        Font fontChinese = new Font(bfChinese, size, style, color);
                        if (familyName != null)
                            fontChinese.setFamily(familyName);
                        return fontChinese;
                    } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                    }
                }
            });*/
            PdfConverter.getInstance().convert(doc, targetStream, options);
//            File file = new File(tmp);
//            file.delete();  //刪除word文件
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(doc);
            IOUtils.closeQuietly(targetStream);
            IOUtils.closeQuietly(sourceStream);
        }
    }

    public static int getPdfPage(String filepath){
        int pagecount = 0;
        PdfReader reader;
        try {
            reader = new PdfReader(filepath);
            pagecount= reader.getNumberOfPages();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("pdf的总页数为:" + pagecount);
        return pagecount;
    }

    private static void showBookmark ( Map bookmark ) {
        System.out.println ( bookmark.get ( "Title" )) ;
        ArrayList kids = (ArrayList) bookmark.get ( "Kids" ) ;
        if ( kids == null )
            return ;
        for (Iterator i = kids.iterator (); i.hasNext () ; ) {
            showBookmark (( Map ) i.next ()) ;
        }
    }

    private static void readPdfPage(String filepath) throws Exception{
        PdfReader reader = new PdfReader ( filepath ) ;
        List list = SimpleBookmark.getBookmark ( reader ) ;
        for ( Iterator i = list.iterator () ; i.hasNext () ; ) {
            showBookmark (( Map ) i.next ()) ;
        }
    }


    public static void main(String[] args) throws Exception{
        long start = System.currentTimeMillis();
        String filepath = "F:\\1\\报告.docx";
        String outpath = "F:\\1\\pdf1.pdf";
//
//        wordConverterToPdf(filepath,outpath,"STSong-Light","UniGB-UCS2-H");
//
//        getPdfPage(outpath);
//        readPdfPage(outpath);
        readPDF(outpath);


       /* InputStream source;
        OutputStream target;
        try {
            source = new FileInputStream(filepath);
            target = new FileOutputStream(outpath);
            Map<String, String> params = new HashMap<String, String>();


            PdfOptions options = PdfOptions.create();

            wordConverterToPdf(source, target, options, params);

        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }*/
        long time = System.currentTimeMillis() - start;
        System.out.println("耗时："+time);
    }

    public static void printBookmarks(PDOutlineNode bookmark, String indentation) throws IOException{
        PDOutlineItem current = bookmark.getFirstChild();
        while(current != null){
            int pages = 0;
            if(current.getDestination() instanceof PDPageDestination){
                PDPageDestination pd = (PDPageDestination) current.getDestination();
                pages = pd.retrievePageNumber() + 1;
            }
            if (current.getAction()  instanceof PDActionGoTo) {
                PDActionGoTo gta = (PDActionGoTo) current.getAction();
                if (gta.getDestination() instanceof PDPageDestination) {
                    PDPageDestination pd = (PDPageDestination) gta.getDestination();
                    pages = pd.retrievePageNumber() + 1;
                }
            }
            if (pages == 0) {
                System.out.println(indentation+current.getTitle());
            }else{
                System.out.println(indentation+current.getTitle()+"  "+pages);
            }
            printBookmarks( current, indentation + "    " );
            current = current.getNextSibling();
        }
    }
    public static void readPDF(String filePath){
        File file = new File(filePath);
        PDDocument doc = null;
        FileInputStream  fis = null;
        try {
            fis = new FileInputStream(file);
            PDFParser parser = new PDFParser(new RandomAccessBuffer(fis));
            parser.parse();
            doc = parser.getPDDocument();
            PDDocumentOutline outline = doc.getDocumentCatalog().getDocumentOutline();
            if (outline != null) {
                printBookmarks(outline, "");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
