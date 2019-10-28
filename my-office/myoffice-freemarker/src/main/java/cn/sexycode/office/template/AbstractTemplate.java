package cn.sexycode.office.template;

import cn.sexycode.office.Config;
import cn.sexycode.util.core.factory.BeanFactoryUtil;
import cn.sexycode.util.core.file.FileUtils;
import cn.sexycode.util.core.file.ZipUtils;
import cn.sexycode.util.core.io.IOUtils;
import cn.sexycode.util.core.xml.XmlUtil;
import freemarker.template.Template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public abstract class AbstractTemplate {
    /**
     * 将模板文件解压，找到实际的模板
     *
     * @param templateFile
     * @return
     * @throws IOException
     */
    protected Template prepare(String templateFile) throws IOException {
        String templateDir = FileUtils.withoutSuffix(templateFile);
        String fileName = FileUtils.fileName(templateDir);
        templateDir = Config.getWorkDir() + File.separator + fileName;
        ZipUtils.unZip(new File(templateFile), templateDir);
        File documentFile = new File(templateDir, getXmlPath());
        String replace = XmlUtil.decode(new String(IOUtils.toByteArray(new FileInputStream(documentFile))));
        new FileOutputStream(documentFile).write(replace.getBytes());
        return BeanFactoryUtil.getBeanFactory().getBean(FreemarkerConfiguration.class).getWorkConfiguration()
                .getTemplate(fileName + File.separator + getXmlPath());
    }

    protected String getXmlPath() {
        return WordTemplate.DOC_TEXT;
    }
}
