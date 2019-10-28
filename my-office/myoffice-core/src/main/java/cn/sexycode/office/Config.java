package cn.sexycode.office;

public class Config {
    private static String workDir = System.getProperty("java.io.tmpdir") + "solof";

    /**
     * 单元格处理器是否忽略行数据，这个开关涉及到读取模式：为true时，会在内部先读一行的数据之后再调用单元格处理器；为false时，读取到每个单元格时，都会调用单元格处理器
     */
    public static boolean cellHandlerSkipRowData = false;

    public static String getWorkDir() {
        return workDir;
    }

    public static void setWorkDir(String workDir) {
        Config.workDir = workDir;
    }
}
