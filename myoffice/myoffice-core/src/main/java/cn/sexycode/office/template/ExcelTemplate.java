package cn.sexycode.office.template;

/**
 * excel
 *
 * @author qinzaizhen
 */
public interface ExcelTemplate extends Template {
    /**
     * 要替换的sharedStrings.xml的位置
     */
    String SHARED_STRINGS = "xl/sharedStrings.xml";

    String EXCEL_SUFFIX = ".xlsx";

}
