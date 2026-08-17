package io.exceltools.model;

import io.exceltools.style.ExcelStyleProvider;

/**
 * Excel 导出配置模型。
 *
 * <p>用于灵活控制单次导出行为。所有字段均为可选（null 表示「不覆盖」），
 * 导出时会与 {@code application.yml} 中 {@code excel.export.*} 的全局默认值做合并：
 * 配置中已显式设置的值优先，未设置的项回落到全局默认值。</p>
 *
 * <p>推荐通过 {@link #builder()} 链式构建，例如：</p>
 * <pre>{@code
 * ExcelExportConfig config = ExcelExportConfig.builder()
 *         .sheetName("用户列表")
 *         .title("2026年用户数据报表")
 *         .freezeHeader(true)
 *         .build();
 * }</pre>
 *
 * @author exceltools
 * @since 1.0.0
 */
public class ExcelExportConfig {

    /**
     * 导出文件名（不含扩展名时自动追加 .xlsx）。
     * 为空时取全局配置 {@code excel.export.default-file-name}。
     */
    private String fileName;

    /**
     * Sheet 名称。为空时取全局配置 {@code excel.export.default-sheet-name}。
     * 多 Sheet 导出时，各 {@link SheetData} 的 sheetName 优先于该值。
     */
    private String sheetName;

    /**
     * 标题行文字。设置后会在表头上方输出一行合并的标题（使用标题样式）。
     * 为 null 时不输出标题行。
     */
    private String title;

    /**
     * 表头是否冻结（滚动时固定首行/前两行）。为 null 时取全局配置 {@code excel.export.freeze-header}。
     */
    private Boolean freezeHeader;

    /**
     * 是否自动计算列宽。为 null 时取全局配置 {@code excel.export.auto-column-width}。
     */
    private Boolean autoColumnWidth;

    /**
     * 分批导出时每批拉取的数据条数。为 null 时取全局配置 {@code excel.export.batch-size}。
     */
    private Integer batchSize;

    /**
     * 大数据量流式导出（SXSSF）时内存中保留的行数（窗口大小）。
     * 为 null 时取全局配置 {@code excel.export.window-rows}。
     */
    private Integer windowRows;

    /**
     * 布尔值 true 的展示文本。为 null 时取全局配置 {@code excel.export.boolean-true-text}。
     */
    private String booleanTrueText;

    /**
     * 布尔值 false 的展示文本。为 null 时取全局配置 {@code excel.export.boolean-false-text}。
     */
    private String booleanFalseText;

    /**
     * 数据行类的显式声明。当导出的集合为空（无法从首元素推断类型）时，
     * 用于解析表头结构；亦可用于强制指定解析类。
     */
    private Class<?> dataClass;

    /**
     * 自定义样式提供者。为 null 时使用组件内置默认样式
     * （默认实现见 {@link io.exceltools.style.DefaultStyleProvider}）。
     */
    private ExcelStyleProvider styleProvider;

    /**
     * 私有构造器，请使用 {@link #builder()} 创建实例。
     */
    private ExcelExportConfig() {
    }

    /**
     * 创建配置构建器。
     *
     * @return 配置构建器实例
     */
    public static Builder builder() {
        return new Builder();
    }

    public String getFileName() {
        return fileName;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getTitle() {
        return title;
    }

    public Boolean getFreezeHeader() {
        return freezeHeader;
    }

    public Boolean getAutoColumnWidth() {
        return autoColumnWidth;
    }

    public Integer getBatchSize() {
        return batchSize;
    }

    public Integer getWindowRows() {
        return windowRows;
    }

    public String getBooleanTrueText() {
        return booleanTrueText;
    }

    public String getBooleanFalseText() {
        return booleanFalseText;
    }

    public Class<?> getDataClass() {
        return dataClass;
    }

    public ExcelStyleProvider getStyleProvider() {
        return styleProvider;
    }

    /**
     * ExcelExportConfig 构建器。
     */
    public static final class Builder {

        private final ExcelExportConfig config = new ExcelExportConfig();

        private Builder() {
        }

        /**
         * 设置导出文件名（可含路径；不含扩展名时自动追加 .xlsx）。
         *
         * @param fileName 文件名
         * @return 构建器自身
         */
        public Builder fileName(String fileName) {
            config.fileName = fileName;
            return this;
        }

        /**
         * 设置 Sheet 名称。
         *
         * @param sheetName Sheet 名
         * @return 构建器自身
         */
        public Builder sheetName(String sheetName) {
            config.sheetName = sheetName;
            return this;
        }

        /**
         * 设置标题行文字（置于表头上方，合并单元格 + 标题样式）。
         *
         * @param title 标题文字，null 表示不输出标题行
         * @return 构建器自身
         */
        public Builder title(String title) {
            config.title = title;
            return this;
        }

        /**
         * 设置是否冻结表头。
         *
         * @param freezeHeader true 冻结
         * @return 构建器自身
         */
        public Builder freezeHeader(Boolean freezeHeader) {
            config.freezeHeader = freezeHeader;
            return this;
        }

        /**
         * 设置是否自动计算列宽。
         *
         * @param autoColumnWidth true 自动计算
         * @return 构建器自身
         */
        public Builder autoColumnWidth(Boolean autoColumnWidth) {
            config.autoColumnWidth = autoColumnWidth;
            return this;
        }

        /**
         * 设置分批导出每批拉取条数。
         *
         * @param batchSize 每批条数
         * @return 构建器自身
         */
        public Builder batchSize(Integer batchSize) {
            config.batchSize = batchSize;
            return this;
        }

        /**
         * 设置流式导出内存窗口行数。
         *
         * @param windowRows 窗口行数
         * @return 构建器自身
         */
        public Builder windowRows(Integer windowRows) {
            config.windowRows = windowRows;
            return this;
        }

        /**
         * 设置布尔值 true 的展示文本（默认「是」）。
         *
         * @param text 展示文本
         * @return 构建器自身
         */
        public Builder booleanTrueText(String text) {
            config.booleanTrueText = text;
            return this;
        }

        /**
         * 设置布尔值 false 的展示文本（默认「否」）。
         *
         * @param text 展示文本
         * @return 构建器自身
         */
        public Builder booleanFalseText(String text) {
            config.booleanFalseText = text;
            return this;
        }

        /**
         * 显式指定数据行类（空数据时用于解析表头结构）。
         *
         * @param dataClass 数据类
         * @return 构建器自身
         */
        public Builder dataClass(Class<?> dataClass) {
            config.dataClass = dataClass;
            return this;
        }

        /**
         * 设置自定义样式提供者（标题/表头/数据样式）。
         *
         * @param styleProvider 样式提供者
         * @return 构建器自身
         */
        public Builder styleProvider(ExcelStyleProvider styleProvider) {
            config.styleProvider = styleProvider;
            return this;
        }

        /**
         * 构建配置对象。
         *
         * @return 配置实例
         */
        public ExcelExportConfig build() {
            return config;
        }
    }
}
