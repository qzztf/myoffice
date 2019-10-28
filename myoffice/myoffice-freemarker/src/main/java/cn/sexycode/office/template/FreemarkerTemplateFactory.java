package cn.sexycode.office.template;

import cn.sexycode.util.core.factory.BeanFactory;
import cn.sexycode.util.core.factory.BeanFactoryUtil;
import cn.sexycode.util.core.file.FileUtils;

import java.util.HashMap;
import java.util.Map;

/**
 * @author qinzaizhen
 */
public class FreemarkerTemplateFactory extends AbstractTemplateFactory {
    @Override
    public Template findTemplate(String templateName) {
        if (WordTemplate.WORD_SUFFIX.equalsIgnoreCase(FileUtils.toSuffix(templateName))) {
            return new FreemarkerWordTemplate(templateName);
        }
        if (ExcelTemplate.EXCEL_SUFFIX.equalsIgnoreCase(FileUtils.toSuffix(templateName))) {
            return new FreemarkerExcelTemplate(templateName);
        }
        throw new TemplateNotFoundException("无法找到[" + templateName + "]对应的模板");
    }

    @Override
    public void initConfig() {
        BeanFactoryUtil.setBeanFactory(new BeanFactory() {
            private Map<Class, Object> objectMap = new HashMap<Class, Object>(1) {{
                put(FreemarkerConfiguration.class, new FreemarkerConfiguration());
            }};

            @Override
            public <T> T getBean(Class<T> clazz) {
                return (T) objectMap.get(clazz);
            }
        });
    }
}
