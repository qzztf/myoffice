package cn.sexycode.office.reader.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * @author qinzaizhen
 */
public abstract class SimpleExtractor<T> extends AbstractExtractor<T> {
    private static final Logger LOGGER = LoggerFactory.getLogger(SimpleExtractor.class);

    private List<T> datas = new ArrayList<>();

    @Override
    public T read(String labelRow, int rowNum, List<CellData> rowData) {
        T data = super.read(labelRow, rowNum, rowData);
        datas.add(data);
        return data;
    }

    public List<T> getData(String filePath) {
        DefaultExcelReader defaultExcelReader = new DefaultExcelReader();
        defaultExcelReader.addRowHandler(this);
        try {
            defaultExcelReader.read(Files.newInputStream(Paths.get(filePath)));
        } catch (IOException e) {
            LOGGER.error("读取文件发生异常", e);
        }
        return datas;
    }
}
