package cn.sexycode.office.reader.excel;

import cn.sexycode.office.reader.ParseException;
import cn.sexycode.util.core.cls.TypeResolver;
import cn.sexycode.util.core.cls.internal.JavaReflectionManager;
import cn.sexycode.util.core.object.ReflectHelper;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import static cn.sexycode.util.core.cls.XClass.ACCESS_FIELD;

/**
 * @author qinzaizhen
 */
public abstract class AbstractExtractor<T> implements RowHandler<T> {

    private Map<String, String> mapping;

    private Class<?> modelClass;

    protected AbstractExtractor() {
        initMapping();
    }

    class Mapping {
        private String labelCol;

        private String field;

        Mapping(String labelCol, String field) {
            this.labelCol = labelCol;
            this.field = field;
        }

        String getLabelCol() {
            return labelCol;
        }

        String getField() {
            return field;
        }
    }

    private void initMapping() {
        Class<?>[] rawArguments = TypeResolver.resolveRawArguments(AbstractExtractor.class, getClass());
        modelClass = rawArguments[0];
        if (Objects.isNull(modelClass)) {
            throw new IllegalStateException("无法确定model类型");
        }
        JavaReflectionManager javaReflectionManager = new JavaReflectionManager();
        mapping = javaReflectionManager.toXClass(modelClass).getDeclaredProperties(ACCESS_FIELD).stream()
                .map(xProperty -> new Mapping(xProperty.getAnnotation(CellField.class).labelCol(), xProperty.getName()))
                .collect(Collectors.toMap(Mapping::getLabelCol, Mapping::getField));
        System.out.println(mapping);
    }

    @Override
    public T read(String labelRow, int rowNum, List<CellData> rowData) {
        Object model;
        try {
            model = Objects.requireNonNull(ReflectHelper.getDefaultConstructor(modelClass)).newInstance();
        } catch (Exception e) {
            throw new ParseException("初始化异常", e);
        }
        Object finalModel = model;
        rowData.stream().filter(data -> mapping.containsKey(data.getLabelCol())).peek(data -> {
            String field = mapping.get(data.getLabelCol());
            try {
                ReflectHelper.findSetterMethod(modelClass, field, null).invoke(finalModel, data.getData());
            } catch (Exception e) {
                throw new ParseException("赋值异常", e);
            }
        }).collect(Collectors.toList());
        return (T) model;
    }
}
