import cn.sexycode.office.reader.excel.CellField;
import cn.sexycode.office.reader.excel.DefaultExcelReader;
import cn.sexycode.office.reader.excel.SimpleExtractor;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;

@RunWith(JUnit4.class)
public class ExcelReaderTest {
    @Test
    public void testRead() throws IOException {
        DefaultExcelReader defaultExcelReader = new DefaultExcelReader();
        defaultExcelReader.addCellHandler((rowData, data) -> {
            System.out.println("labelCol ----> " + data.getLabelRow());
            System.out.println("labelY ----> " + data.getLabelCol());
            System.out.println("rowNum ----> " + data.getRowNum());
            System.out.println("colNum ----> " + data.getColNum());
            System.out.println("rowData ----> " + rowData);
            System.out.println("data ----> " + data);
            return null;
        });
        defaultExcelReader.read(Files.newInputStream(Paths.get("/Users/qzz/业务追踪表1557385217666.xlsx")));
    }

    public static class Model {
        @CellField(labelCol = "A")
        private String name;

        @Override
        public String toString() {
            return "Model{" + "name='" + name + '\'' + '}';
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    @Test
    public void extractor() {

        List<Model> collect = new SimpleExtractor<Model>() {
        }.getData("e:/x5/业务追踪表1557384063836.xlsx").stream().peek(System.out::println).collect(Collectors.toList());
    }
}
