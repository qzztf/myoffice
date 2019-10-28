package cn.sexycode.office.reader.excel;

import cn.sexycode.office.reader.Reader;

import java.util.List;
import java.util.Optional;

/**
 * @author qzz
 */
public interface ExcelReader extends Reader {
    /**
     * 单元格处理器
     *
     * @param cellHandlers
     */
    void setCellHandlers(List<CellHandler> cellHandlers);

    /**
     * 返回阅读器的行处理器
     *
     * @return List<CellHandler>
     */
    List<CellHandler> getCellHandlers();

    /**
     * 添加处理器
     *
     * @param cellHandler
     */
    default void addCellHandler(CellHandler cellHandler) {
        Optional.ofNullable(getCellHandlers()).ifPresent(cellHandlers -> cellHandlers.add(cellHandler));
    }

    /**
     * 行处理器
     *
     * @param rowHandlers
     */
    void setRowHandlers(List<RowHandler> rowHandlers);

    /**
     * 返回行处理器
     *
     * @return
     */
    List<RowHandler> getRowHandlers();

    /**
     * 添加行处理器
     *
     * @param rowHandler
     */
    default void addRowHandler(RowHandler rowHandler) {
        Optional.ofNullable(getRowHandlers()).ifPresent(rowHandlers -> rowHandlers.add(rowHandler));
    }

    /**
     * 表格处理器
     *
     * @param sheetHandlers
     */
    void setSheetHandlers(List<SheetHandler> sheetHandlers);

    /**
     * @return 返回表格处理器
     */
    List<SheetHandler> getSheetHandlers();

    /**
     * 添加表格处理器
     *
     * @param sheetHandler
     */
    default void addSheetHandler(SheetHandler sheetHandler) {
        Optional.ofNullable(getSheetHandlers()).ifPresent(sheetHandlers -> sheetHandlers.add(sheetHandler));
    }
}
