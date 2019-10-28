package cn.sexycode.office.reader.excel;

import cn.sexycode.util.core.lang.Order;

import java.util.List;

/**
 * 单行读取的处理器
 *
 * @author qzz
 */
public interface RowHandler<T> extends Order {
    /**
     * @param labelRow 横坐标
     * @param rowNum
     * @param rowData  行数据
     */
    T read(String labelRow, int rowNum, List<CellData> rowData);
}
