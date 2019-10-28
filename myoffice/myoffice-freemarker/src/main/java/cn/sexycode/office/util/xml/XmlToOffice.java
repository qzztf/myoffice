package cn.sexycode.office.util.xml;

import cn.sexycode.office.template.ExcelTemplate;
import cn.sexycode.office.template.WordTemplate;
import cn.sexycode.util.core.file.ZipUtils;
import freemarker.template.Template;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * @author qzz
 */
public class XmlToOffice {
    private static final Logger LOGGER = LoggerFactory.getLogger(XmlToOffice.class);

    /**
     * 根据数据生成文本
     *
     * @param dataMap     数据
     * @param outFilePath 生成的document.xml和document.xml.rels对应的目录名称
     * @param template    模板对象
     * @throws Exception
     */
    private static void toText(Map<String, Object> dataMap, String outFilePath, Template template) throws Exception {
        File docFile = new File(outFilePath);
        FileOutputStream fos = new FileOutputStream(docFile);
        Writer out = new BufferedWriter(new OutputStreamWriter(fos), 10240);
        template.process(dataMap, out);
        out.close();
    }

    /**
     * 生成word docx
     *
     * @param dataMap         数据
     * @param template        模板
     * @param destDocFilePath 生成的document.xml和document.xml.rels对应的目录名称
     * @throws Exception
     */
    public static void makeWord(Map<String, Object> dataMap, Template template, String destDocFilePath,
            String sourceDocFilePath) throws Exception {

        try {
            String outFilePath = sourceDocFilePath + ".xml";
            toText(dataMap, outFilePath, template);

     /*   template = configuration.getTemplate(WordTemplate.IMG_NAME);
        outFilePath = outDocFilePath + ".xml.rels";
        toText(dataMap, outFilePath, template);*/

            //替换原word中的xml文件
            replace(destDocFilePath, sourceDocFilePath, outFilePath, WordTemplate.DOC_TEXT);
        } catch (Exception e) {
            LOGGER.error("异常", e);
            throw e;
        }
    }

    private static void replace(String outDocFilePath, String sourceDocFilePath, String outFilePath, String docText)
            throws IOException {
        ZipInputStream zipInputStream = ZipUtils.wrapZipInputStream(new FileInputStream(new File(sourceDocFilePath)));
        ZipOutputStream zipOutputStream = ZipUtils.wrapZipOutputStream(new FileOutputStream(new File(outDocFilePath)));
        //            File fileImg = new File(outDocFilePath + ".xml.rels");

        Map<String, InputStream> map = new HashMap<>(1);
        map.put(docText, new FileInputStream(outFilePath));
        ZipUtils.replaceItem(map, zipInputStream, zipOutputStream);
    }

    /**
     * 生成word docx
     *
     * @param dataMap        数据
     * @param outDocFilePath 生成的document.xml和document.xml.rels对应的目录名称
     * @throws Exception
     */
    public static void makeExcel(Map<String, Object> dataMap, Template template, String outDocFilePath,
            String sourceDocFilePath) throws Exception {

        try {
            String outFilePath = sourceDocFilePath + ".xml";
            toText(dataMap, outFilePath, template);
            //替换原xlsx中的xml文件
            replace(outDocFilePath, sourceDocFilePath, outFilePath, ExcelTemplate.SHARED_STRINGS);

        } catch (Exception e) {
            LOGGER.error("异常", e);
            throw e;
        }
    }

    /*  *//**
     * 生成pdf
     *//*
    public static  void makePdfByXcode(String ftlPath,String docFilePath){
        try {
            XWPFDocument document=new XWPFDocument(new FileInputStream(new File(docFilePath+".docx")));
            File outFile=new File(docFilePath+".pdf");
                if(!outFile.getParentFile().exists()){
                    outFile.getParentFile().mkdirs();
            }
            OutputStream out=new FileOutputStream(outFile);
            PdfOptions options= PdfOptions.getDefault();
            IFontProvider iFontProvider = new IFontProvider() {
                @Override
                public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                    try {
                        BaseFont bfChinese = null;
                        if( OS.indexOf("linux")>=0){
                            bfChinese =  BaseFont.createFont(ftlPath+"/font/msyh.ttf", BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);
                        }else{
                            bfChinese =  BaseFont.createFont("C:/WINDOWS/Fonts/STSONG.TTF", BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);
                        }
                        Font fontChinese = new Font(bfChinese, size, style, color);
                        if (familyName != null)
                            fontChinese.setFamily(familyName);
                        return fontChinese;
                    } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                    }
                }
            };
            options.fontProvider( iFontProvider );
            PdfConverter.getInstance().convert(document,out,options);

        }
        catch (  Exception e) {
            e.printStackTrace();
        }
    }

*/
}