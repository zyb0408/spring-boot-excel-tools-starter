package io.exceltools.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * Excel 导出全局配置属性类。
 *
 * <p>对应 {@code application.yml} 中的 {@code excel.export.*} 配置项，
 * 由自动配置通过 {@code @EnableConfigurationProperties} 绑定并注入导出服务，
 * 作为所有导出的全局默认值。方法级 {@code ExcelExportConfig} 可覆盖其中任意一项。</p>
 *
 * <h3>配置示例：</h3>
 * <pre>{@code
 * excel:
 *   export:
 *     default-file-name: 导出数据.xlsx   # 默认导出文件名
 *     default-sheet-name: Sheet1         # 默认 Sheet 名
 *     default-title:                     # 默认标题行（留空则不输出标题行）
 *     auto-column-width: true            # 是否自动计算列宽
 *     freeze-header: true                # 是否冻结表头
 *     batch-size: 10000                  # 分批导出每批拉取条数
 *     window-rows: 100                   # 流式导出内存窗口行数
 *     boolean-true-text: 是              # 布尔 true 展示文本
 *     boolean-false-text: 否             # 布尔 false 展示文本
 * }</pre>
 *
 * @author exceltools
 * @since 1.0.0
 */
@ConfigurationProperties(prefix = "excel.export")
public class ExcelExportProperties {

    /**
     * 默认导出文件名。
     * <p>方法未指定文件名时使用；可含目录，不含扩展名时自动追加 .xlsx。</p>
     */
    private String defaultFileName = "导出数据.xlsx";

    /**
     * 默认 Sheet 名称。
     * <p>方法未指定 Sheet 名时使用。</p>
     */
    private String defaultSheetName = "Sheet1";

    /**
     * 默认标题行文字。
     * <p>设置后导出会在表头上方输出一行合并标题；留空（null）表示不输出标题行。</p>
     */
    private String defaultTitle;

    /**
     * 是否自动计算列宽。
     * <p>按表头与数据内容估算列宽（中文按 2 字符宽计），受 Excel 单列 255 字符上限约束。</p>
     */
    private boolean autoColumnWidth = true;

    /**
     * 是否冻结表头。
     * <p>滚动时表头行（及标题行）保持固定显示。</p>
     */
    private boolean freezeHeader = true;

    /**
     * 分批导出每批拉取的数据条数。
     * <p>用于 {@code exportLargeByBatch} 系列方法，控制每次回调 {@code BatchDataProvider} 拉取的行数。</p>
     */
    private int batchSize = 10000;

    /**
     * 流式导出（SXSSF）内存窗口行数。
     * <p>内存中保留的最近写入行数，超出部分自动落盘到临时文件，内存占用与总数据量无关。</p>
     */
    private int windowRows = 100;

    /**
     * 布尔值 true 的展示文本（默认「是」）。
     */
    private String booleanTrueText = "是";

    /**
     * 布尔值 false 的展示文本（默认「否」）。
     */
    private String booleanFalseText = "否";

    public String getDefaultFileName() {
        return defaultFileName;
    }

    public void setDefaultFileName(String defaultFileName) {
        this.defaultFileName = defaultFileName;
    }

    public String getDefaultSheetName() {
        return defaultSheetName;
    }

    public void setDefaultSheetName(String defaultSheetName) {
        this.defaultSheetName = defaultSheetName;
    }

    public String getDefaultTitle() {
        return defaultTitle;
    }

    public void setDefaultTitle(String defaultTitle) {
        this.defaultTitle = defaultTitle;
    }

    public boolean isAutoColumnWidth() {
        return autoColumnWidth;
    }

    public void setAutoColumnWidth(boolean autoColumnWidth) {
        this.autoColumnWidth = autoColumnWidth;
    }

    public boolean isFreezeHeader() {
        return freezeHeader;
    }

    public void setFreezeHeader(boolean freezeHeader) {
        this.freezeHeader = freezeHeader;
    }

    public int getBatchSize() {
        return batchSize;
    }

    public void setBatchSize(int batchSize) {
        this.batchSize = batchSize;
    }

    public int getWindowRows() {
        return windowRows;
    }

    public void setWindowRows(int windowRows) {
        this.windowRows = windowRows;
    }

    public String getBooleanTrueText() {
        return booleanTrueText;
    }

    public void setBooleanTrueText(String booleanTrueText) {
        this.booleanTrueText = booleanTrueText;
    }

    public String getBooleanFalseText() {
        return booleanFalseText;
    }

    public void setBooleanFalseText(String booleanFalseText) {
        this.booleanFalseText = booleanFalseText;
    }
}
