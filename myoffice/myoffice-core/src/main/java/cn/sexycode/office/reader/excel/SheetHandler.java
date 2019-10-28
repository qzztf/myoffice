package cn.sexycode.office.reader.excel;

import cn.sexycode.util.core.lang.Order;

/**
 * 表格处理器
 *
 * @author qzz
 */
public interface SheetHandler extends Order {
    void read(int sheetIndex, String title);
}
