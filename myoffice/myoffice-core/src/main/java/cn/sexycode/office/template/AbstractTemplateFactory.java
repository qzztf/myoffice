package cn.sexycode.office.template;

/**
 * @author qinzaizhen
 */
public class AbstractTemplateFactory implements TemplateFactory {
    protected AbstractTemplateFactory() {
        initConfig();
    }

    @Override
    public Template findTemplate(String templateName) {
        throw new TemplateNotFoundException("请实现此方法");
    }

    @Override
    public void initConfig() {

    }
}
