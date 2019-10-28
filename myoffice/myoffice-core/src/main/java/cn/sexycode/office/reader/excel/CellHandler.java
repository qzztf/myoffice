package cn.sexycode.office.reader.excel;

import cn.sexycode.util.core.lang.Order;

import java.util.List;

/**
 * 单元格读取处理器
 *
 * @author qzz
 */
public interface CellHandler<T> extends Order {
    /**
     * @param rowData 行数据 不要修改list的长度，内部有遍历
     * @param data    原始数据
     */
    T read(List<CellData> rowData, CellData data);
}
