package cn.sexycode.office.template;

import java.util.Iterator;
import java.util.Objects;
import java.util.ServiceLoader;

public interface TemplateFactory {

    static TemplateFactory loadTemplateFactory() {

        Iterator<TemplateFactory> iterator = ServiceLoader.load(TemplateFactory.class).iterator();
        TemplateFactory templateFactory = null;
        if (iterator.hasNext()) {
            templateFactory = iterator.next();
        }

        if (iterator.hasNext()) {
            throw new IllegalStateException("Loaded more than one TemplateFactory");
        }

        if (Objects.isNull(templateFactory)) {
            throw new IllegalStateException("Can't find one TemplateFactory");
        }
        return templateFactory;
    }

    /**
     * 根据模板名找到对应的模板
     *
     * @param templateName 模板名
     * @return Template
     */
    Template findTemplate(String templateName);

    void initConfig();

}
