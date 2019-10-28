package cn.sexycode.office.template;

import java.util.Map;

/**
 * @author qinzaizhen
 */
public interface Template {
    /**
     * 处理模板，转换成对应的文件
     *
     * @param outFile
     * @param dataModel
     */
    void process(String outFile, Map<String, Object> dataModel);
}
