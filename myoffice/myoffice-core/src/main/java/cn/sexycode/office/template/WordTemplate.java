package cn.sexycode.office.template;

/**
 * word
 *
 * @author qinzaizhen
 */
public interface WordTemplate extends Template {
    /**
     * 要替换的document.xml的位置
     */
    String DOC_TEXT = "word/document.xml";
    /**
     * 要替换的document.xml.rels的位置
     */
    String IMG_NAME = "word/_rels/document.xml.rels";
    /**
     * 图片meida的位置
     */
    String MEDIA_NAME = "word/media/";

    String WORD_SUFFIX = ".docx";

}
