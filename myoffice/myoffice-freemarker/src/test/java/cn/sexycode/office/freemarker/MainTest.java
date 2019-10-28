package cn.sexycode.office.freemarker;

import cn.sexycode.office.template.FreemarkerTemplateFactory;
import cn.sexycode.office.template.TemplateFactory;
import cn.sexycode.office.template.TemplateUtil;
import cn.sexycode.util.core.file.ZipUtils;
import cn.sexycode.util.core.io.IOUtils;
import cn.sexycode.util.core.str.StringHelper;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

@RunWith(JUnit4.class)
public class MainTest {
    @Test
    public void test() throws Exception {
        //        new WordUtil().replaceText(new XWPFDocument(new FileInputStream("e:/x5/鼎和房屋抵押合同履约保证保险投保单.docx")));

        /*Config cfg = new Config(Config.VERSION_2_3_28);
        cfg.setDirectoryForTemplateLoading(new File("e:/x5"));
        Template template =cfg.getTemplate("鼎和房屋抵押合同履约保证保险投保单.docx");*/
        HashMap<String, Object> dataModel = new HashMap<>();
        dataModel.put("city", 11111);
        dataModel.put("text", 11111);
        dataModel.put("lists", new ArrayList<>(Arrays.asList("11111", "22222")));
        dataModel.put("bankName", 11111);
        dataModel.put("creditCode", 11111);
        dataModel.put("bankContactPhone", 11111);
        dataModel.put("bankAddress", 11111);
        //        template.process(dataModel, new OutputStreamWriter(new FileOutputStream("e:/x5/ffff.docx")));

        ZipUtils.unZip(new File("e:/template/demot.docx"), "e:/template/demot");
        String template = new String(IOUtils.toByteArray(new FileInputStream("e:/template/demot/word/document.xml")));
        String replace = StringHelper.replace(template, "&lt;", "<").replace("&gt;", ">");
        new FileOutputStream("e:/template/demot/word/document.xml").write(replace.getBytes());
        //        XmlToOffice.makeWord(dataModel, "e:/template/demot/word", "e:/template/demo", "e:/template/demot.docx");

    }

    @Test
    public void testTemplate() {
        HashMap<String, Object> dataModel = new HashMap<>();
        dataModel.put("city", 11111);
        dataModel.put("text", 11111);
        dataModel.put("lists", new ArrayList<>(Arrays.asList("11111", "22222")));
        dataModel.put("bankName", 11111);
        dataModel.put("creditCode", 11111);
        dataModel.put("bankContactPhone", 11111);
        dataModel.put("bankAddress", 11111);
        (TemplateUtil.getTemplate("e:/template/demot.docx")).process("e:/template/demo.docx", dataModel);
        (TemplateUtil.getTemplate("e:/template/exchangeInfo1.xlsx"))
                .process("e:/template/exchangeInfo.xlsx", dataModel);
    }

    @Test
    public void testTemplateFactory() {
        HashMap<String, Object> dataModel = new HashMap<>();
        dataModel.put("city", 11111);
        dataModel.put("text", 11111);
        dataModel.put("lists", new ArrayList<>(Arrays.asList("11111", "22222")));
        dataModel.put("bankName", 11111);
        dataModel.put("creditCode", 11111);
        dataModel.put("bankContactPhone", 11111);
        dataModel.put("bankAddress", 11111);
        TemplateFactory factory = new FreemarkerTemplateFactory();
        factory.findTemplate("e:/template/demot.docx").process("e:/template/demo.docx", dataModel);
        /*(TemplateUtil.getTemplate("e:/template/exchangeInfo1.xlsx"))
                .process("e:/template/exchangeInfo.xlsx", dataModel);*/
    }
}
