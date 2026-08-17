package io.exceltools.model;

/**
 * 多 Sheet 导出时的单个 Sheet 数据模型。
 *
 * <p>每个 {@code SheetData} 描述一个 Sheet 的名称、数据与独立配置，
 * 通过 {@link ExcelExportService} 的多 Sheet 导出方法一次生成多个 Sheet。</p>
 *
 * <h3>示例：</h3>
 * <pre>{@code
 * List<SheetData> sheets = Arrays.asList(
 *         SheetData.builder().sheetName("用户").data(userList).build(),
 *         SheetData.builder().sheetName("订单").data(orderList)
 *                 .config(ExcelExportConfig.builder().title("订单报表").build())
 *                 .build()
 * );
 * }</pre>
 *
 * @author exceltools
 * @since 1.0.0
 */
public class SheetData {

    /**
     * Sheet 名称。为空时回落到全局配置或 {@link ExcelExportConfig#getSheetName()}。
     */
    private String sheetName;

    /**
     * 该 Sheet 导出的数据行集合（可为空集合）。
     */
    private java.util.List<?> data;

    /**
     * 该 Sheet 的数据行类。为空且数据集合非空时自动取首元素类型；
     * 数据为空时需显式指定以生成表头。
     */
    private Class<?> dataClass;

    /**
     * 该 Sheet 的独立导出配置，可覆盖全局配置。
     */
    private ExcelExportConfig config;

    /**
     * 私有构造器，请使用 {@link #builder()} 创建实例。
     */
    private SheetData() {
    }

    /**
     * 创建 SheetData 构建器。
     *
     * @return 构建器实例
     */
    public static Builder builder() {
        return new Builder();
    }

    public String getSheetName() {
        return sheetName;
    }

    public java.util.List<?> getData() {
        return data;
    }

    public Class<?> getDataClass() {
        return dataClass;
    }

    public ExcelExportConfig getConfig() {
        return config;
    }

    /**
     * SheetData 构建器。
     */
    public static final class Builder {

        private final SheetData sheetData = new SheetData();

        private Builder() {
        }

        /**
         * 设置 Sheet 名称。
         *
         * @param sheetName Sheet 名
         * @return 构建器自身
         */
        public Builder sheetName(String sheetName) {
            sheetData.sheetName = sheetName;
            return this;
        }

        /**
         * 设置该 Sheet 导出的数据行集合。
         *
         * @param data 数据集合
         * @return 构建器自身
         */
        public Builder data(java.util.List<?> data) {
            sheetData.data = data;
            return this;
        }

        /**
         * 显式指定该 Sheet 的数据行类（数据为空时用于生成表头）。
         *
         * @param dataClass 数据类
         * @return 构建器自身
         */
        public Builder dataClass(Class<?> dataClass) {
            sheetData.dataClass = dataClass;
            return this;
        }

        /**
         * 设置该 Sheet 的独立导出配置。
         *
         * @param config 导出配置
         * @return 构建器自身
         */
        public Builder config(ExcelExportConfig config) {
            sheetData.config = config;
            return this;
        }

        /**
         * 构建 SheetData 实例。
         *
         * @return SheetData 实例
         */
        public SheetData build() {
            return sheetData;
        }
    }
}
