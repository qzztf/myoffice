package cn.sexycode.office.template;

import cn.sexycode.office.util.xml.XmlToOffice;
import freemarker.template.Template;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

/**
 * excel 模板
 *
 * @author qinzaizhen
 */
public class FreemarkerExcelTemplate extends AbstractTemplate implements ExcelTemplate {

    private static final Logger LOGGER = LoggerFactory.getLogger(FreemarkerExcelTemplate.class);

    private String path;

    private String templateName;

    /**
     * 如果是文件，则传文件的完整路径
     *
     * @param path
     */
    public FreemarkerExcelTemplate(String path) {
        this.path = path;
    }

    @Override
    public void process(String outFile, Map<String, Object> dataModel) {

        try {
            Template template = prepare(path);
            XmlToOffice.makeExcel(dataModel, template, outFile, path);
        } catch (Exception e) {
            LOGGER.error("转换异常", e);
        }
    }

    @Override
    protected String getXmlPath() {
        return ExcelTemplate.SHARED_STRINGS;
    }
}
