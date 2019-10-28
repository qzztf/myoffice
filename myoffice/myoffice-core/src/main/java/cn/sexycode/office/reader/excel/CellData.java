package cn.sexycode.office.reader.excel;

import org.apache.poi.ss.util.CellReference;

/**
 * @author qinzaizhen
 */
public class CellData {
    private String labelCol;

    private String labelRow;

    private Object data;

    private int rowNum;

    private int colNum;

    private int sheetIndex;

    public CellData(int sheetIndex, String labelRow, String labelCol, Object data) {
        this.labelCol = labelCol;
        this.labelRow = labelRow;
        this.rowNum = Integer.valueOf(labelRow) - 1;
        this.colNum = CellReference.convertColStringToIndex(labelCol);
        this.data = data;
    }

    public String getLabelCol() {
        return labelCol;
    }

    public String getLabelRow() {
        return labelRow;
    }

    public Object getData() {
        return data;
    }

    public int getRowNum() {
        return rowNum;
    }

    public int getColNum() {
        return colNum;
    }

    @Override
    public String toString() {
        return "CellData{" + "labelCol='" + labelCol + '\'' + ", labelRow='" + labelRow + '\'' + ", data=" + data
                + ", rowNum=" + rowNum + ", colNum=" + colNum + '}';
    }

    public int getSheetIndex() {
        return sheetIndex;
    }
}
