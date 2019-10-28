package cn.sexycode.office.reader;

import java.io.InputStream;

/**
 * 阅读器
 *
 * @author qzz
 */
public interface Reader {
    /**
     * @param in 输入流
     * @throws ParseException 解析失败时抛出此异常
     */
    void read(InputStream in) throws ParseException;
}
