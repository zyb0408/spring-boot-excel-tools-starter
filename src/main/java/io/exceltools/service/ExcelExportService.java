package io.exceltools.service;

import io.exceltools.model.ExcelExportConfig;
import io.exceltools.model.SheetData;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Excel 统一导出服务接口。
 *
 * <p>团队所有成员的 Excel 导出统一走本接口，保证导出行为（列定义、样式、文件格式）一致。
 * 该接口由自动配置注册为 Spring Bean，直接注入使用即可，例如：</p>
 *
 * <pre>{@code
 * @Autowired
 * private ExcelExportService exportService;
 *
 * // 一行代码导出单表
 * exportService.export(userList, "用户数据.xlsx");
 * }</pre>
 *
 * <p>核心能力：</p>
 * <ul>
 *   <li><strong>单表导出</strong>：{@link #export} 系列，适用于常规数据量；</li>
 *   <li><strong>多 Sheet 导出</strong>：{@link #exportMultiSheet} 系列，一次生成多个 Sheet；</li>
 *   <li><strong>大数据量分批导出</strong>：{@link #exportLarge} / {@link #exportLargeByBatch} 系列，
 *       基于 SXSSF 流式写入，内存占用恒定；</li>
 *   <li><strong>输出目标</strong>：文件路径、字节数组（{@code byte[]}）、输出流（{@code OutputStream}）
 *       三种方式自由选择。</li>
 * </ul>
 *
 * <p>列定义通过 {@link io.exceltools.annotation.ExcelColumn} 注解驱动；
 * 行为通过 {@link ExcelExportConfig} 控制；全局默认值通过
 * {@code application.yml} 中 {@code excel.export.*} 配置。</p>
 *
 * @author exceltools
 * @since 1.0.0
 * @see io.exceltools.annotation.ExcelColumn
 * @see ExcelExportConfig
 */
public interface ExcelExportService {

    // ============================================================
    // 一、单 Sheet 导出（常规数据量，内存中一次生成）
    // ============================================================

    /**
     * 单表导出到文件。
     *
     * @param data     数据行集合（元素为带 {@code @ExcelColumn} 注解的实体对象）
     * @param filePath 导出文件路径（可以是完整路径；仅目录时自动追加默认文件名；
     *                 无扩展名时自动追加 .xlsx）
     * @param config   导出配置（可为 null，使用全局默认配置）
     * @throws IOException 文件写入失败时抛出
     */
    void export(List<?> data, String filePath, ExcelExportConfig config) throws IOException;

    /**
     * 单表导出到文件（使用全局默认配置）。
     *
     * @param data     数据行集合
     * @param filePath 导出文件路径
     * @throws IOException 文件写入失败时抛出
     */
    void export(List<?> data, String filePath) throws IOException;

    /**
     * 单表导出到输出流（如 Servlet 输出流，适用于 Web 下载）。
     *
     * @param data   数据行集合
     * @param out    目标输出流（方法内部不关闭该流，由调用方管理）
     * @param config 导出配置（可为 null，使用全局默认配置）
     * @throws IOException 写入失败时抛出
     */
    void export(List<?> data, OutputStream out, ExcelExportConfig config) throws IOException;

    /**
     * 单表导出到输出流（使用全局默认配置）。
     *
     * @param data 数据行集合
     * @param out  目标输出流
     * @throws IOException 写入失败时抛出
     */
    void export(List<?> data, OutputStream out) throws IOException;

    /**
     * 单表导出为字节数组（适合小数据量，如返回给前端）。
     *
     * @param data   数据行集合
     * @param config 导出配置（可为 null）
     * @return xlsx 文件字节数组
     */
    byte[] exportToBytes(List<?> data, ExcelExportConfig config);

    /**
     * 单表导出为字节数组（使用全局默认配置）。
     *
     * @param data 数据行集合
     * @return xlsx 文件字节数组
     */
    byte[] exportToBytes(List<?> data);

    // ============================================================
    // 二、多 Sheet 导出（一次生成多个 Sheet）
    // ============================================================

    /**
     * 多 Sheet 导出到文件。
     *
     * @param sheets       Sheet 列表（每个元素描述一个 Sheet 的名称、数据与独立配置）
     * @param filePath     导出文件路径
     * @param globalConfig 全局配置（可为 null；每个 Sheet 的独立配置优先于全局配置）
     * @throws IOException 文件写入失败时抛出
     */
    void exportMultiSheet(List<SheetData> sheets, String filePath, ExcelExportConfig globalConfig) throws IOException;

    /**
     * 多 Sheet 导出到文件（使用全局默认配置）。
     *
     * @param sheets   Sheet 列表
     * @param filePath 导出文件路径
     * @throws IOException 文件写入失败时抛出
     */
    void exportMultiSheet(List<SheetData> sheets, String filePath) throws IOException;

    /**
     * 多 Sheet 导出到输出流（如 Servlet 输出流，适用于 Web 下载）。
     *
     * @param sheets       Sheet 列表
     * @param out          目标输出流（方法内部不关闭该流）
     * @param globalConfig 全局配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    void exportMultiSheet(List<SheetData> sheets, OutputStream out, ExcelExportConfig globalConfig)
            throws IOException;

    /**
     * 多 Sheet 导出为字节数组。
     *
     * @param sheets       Sheet 列表
     * @param globalConfig 全局配置（可为 null）
     * @return xlsx 文件字节数组
     */
    byte[] exportMultiSheetToBytes(List<SheetData> sheets, ExcelExportConfig globalConfig);

    /**
     * 多 Sheet 导出为字节数组（使用全局默认配置）。
     *
     * @param sheets Sheet 列表
     * @return xlsx 文件字节数组
     */
    byte[] exportMultiSheetToBytes(List<SheetData> sheets);

    // ============================================================
    // 三、大数据量分批导出（SXSSF 流式写入，内存占用恒定）
    // ============================================================

    /**
     * 大数据量导出到文件（流式写入）。
     *
     * <p>适用于单次内存中已加载全部数据（如几十万行）的场景：使用 SXSSF 流式写入，
     * 超出窗口行数的数据被写入临时文件，避免大对象驻留内存。</p>
     *
     * @param data     数据行集合（内存中已有的全部数据）
     * @param filePath 导出文件路径
     * @param config   导出配置（可为 null；{@code windowRows} 控制内存窗口大小）
     * @throws IOException 文件写入失败时抛出
     */
    void exportLarge(List<?> data, String filePath, ExcelExportConfig config) throws IOException;

    /**
     * 大数据量导出到输出流（流式写入）。
     *
     * @param data   数据行集合
     * @param out    目标输出流
     * @param config 导出配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    void exportLarge(List<?> data, OutputStream out, ExcelExportConfig config) throws IOException;

    /**
     * 大数据量导出为字节数组（流式写入）。
     *
     * <p>注意：字节数组本身需驻留内存，超大文件请优先使用文件/输出流方式。</p>
     *
     * @param data   数据行集合
     * @param config 导出配置（可为 null）
     * @return xlsx 文件字节数组
     */
    byte[] exportLargeToBytes(List<?> data, ExcelExportConfig config);

    /**
     * 分批拉取大数据量导出到文件（推荐用于数据库大数据量导出）。
     *
     * <p>通过 {@link BatchDataProvider} 分页从数据源拉取数据，边拉取边写入，
     * 任何时刻内存中仅保留一批数据，适合百万级数据导出。</p>
     *
     * @param provider 分批数据提供者（分页拉取逻辑由调用方实现）
     * @param filePath 导出文件路径
     * @param config   导出配置（可为 null；{@code batchSize} 控制每批拉取条数）
     * @throws IOException 文件写入失败时抛出
     */
    void exportLargeByBatch(BatchDataProvider<?> provider, String filePath, ExcelExportConfig config)
            throws IOException;

    /**
     * 分批拉取大数据量导出到输出流。
     *
     * @param provider 分批数据提供者
     * @param out      目标输出流
     * @param config   导出配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    void exportLargeByBatch(BatchDataProvider<?> provider, OutputStream out, ExcelExportConfig config)
            throws IOException;
}
