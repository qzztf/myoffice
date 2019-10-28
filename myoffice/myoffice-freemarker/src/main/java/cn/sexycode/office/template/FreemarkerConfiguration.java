package cn.sexycode.office.template;

import cn.sexycode.office.Config;
import freemarker.template.Configuration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;

/**
 * @author qzz
 */
public class FreemarkerConfiguration {
    private static final Logger LOGGER = LoggerFactory.getLogger(FreemarkerConfiguration.class);

    /**
     * 会将模板解压到工作目录，实际的模板是从此目录中加载的xml文件
     */
    private Configuration workConfiguration;

    public FreemarkerConfiguration() {
        workConfiguration = new Configuration(Configuration.VERSION_2_3_0);
        try {
            workConfiguration.setDirectoryForTemplateLoading(new File(Config.getWorkDir()));
        } catch (IOException e) {
            LOGGER.error("加载工作目录失败", e);
        }
    }

    public Configuration getWorkConfiguration() {
        return workConfiguration;
    }

}
