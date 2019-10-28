package cn.sexycode.office.template;

/**
 * 通过此方法查找模板
 *
 * @author qinzaizhen
 */
public class TemplateUtil {
    private static TemplateFactory templateFactory;

    static {
        templateFactory = TemplateFactory.loadTemplateFactory();
        templateFactory.initConfig();
    }

    public static Template getTemplate(String templateName) {
        return templateFactory.findTemplate(templateName);
    }
}
