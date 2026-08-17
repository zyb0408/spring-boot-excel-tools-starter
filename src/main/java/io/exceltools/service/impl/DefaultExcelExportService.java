package io.exceltools.service.impl;

import io.exceltools.annotation.ExcelColumn;
import io.exceltools.annotation.ExcelIgnore;
import io.exceltools.config.ExcelExportProperties;
import io.exceltools.model.Align;
import io.exceltools.model.ExcelExportConfig;
import io.exceltools.model.SheetData;
import io.exceltools.service.BatchDataProvider;
import io.exceltools.service.ExcelExportService;
import io.exceltools.style.DefaultStyleProvider;
import io.exceltools.style.ExcelStyleProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 默认 Excel 导出服务实现。
 *
 * <p>基于 Apache POI 实现，提供单表、多 Sheet、大数据量分批三种导出场景，
 * 支持 {@link ExcelColumn} 注解驱动的列定义与三种输出目标（文件 / 字节数组 / 输出流）。</p>
 *
 * <p>实现要点：</p>
 * <ul>
 *   <li>列解析：优先使用 {@code @ExcelColumn} 标注字段（按 order 升序）；
 *       无任何标注时回退为导出全部非静态非 transient 字段（列名取字段名）；</li>
 *   <li>样式体系：标题 / 表头 / 数据三套样式，可通过 {@link ExcelStyleProvider} 自定义；</li>
 *   <li>大数据量导出：{@link SXSSFWorkbook} 流式写入，内存占用与数据总量无关；
 *       {@link #exportLargeByBatch} 支持边查边写，适合百万级数据；</li>
 *   <li>配置合并：方法级配置未设置的项自动回落至全局配置（{@code excel.export.*}）。</li>
 * </ul>
 *
 * @author exceltools
 * @since 1.0.0
 */
public class DefaultExcelExportService implements ExcelExportService {

    /**
     * 默认日期时间格式。
     */
    private static final String DEFAULT_DATETIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

    /**
     * 默认日期格式。
     */
    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";

    /**
     * 参与列宽自动计算的采样行数上限（避免大数据量时逐个计算拖慢导出）。
     */
    private static final int WIDTH_SAMPLE_ROW_LIMIT = 500;

    /**
     * 全局导出配置（来自 application.yml 的 excel.export.*）。
     */
    private final ExcelExportProperties properties;

    /**
     * 构造默认导出服务。
     *
     * @param properties 全局导出配置；为 null 时使用内置默认值（便于脱离 Spring 容器使用）
     */
    public DefaultExcelExportService(ExcelExportProperties properties) {
        this.properties = properties != null ? properties : new ExcelExportProperties();
    }

    // ============================================================
    // 一、单 Sheet 导出
    // ============================================================

    /**
     * 单表导出到文件。
     *
     * @param data     数据行集合
     * @param filePath 导出文件路径
     * @param config   导出配置（可为 null）
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void export(List<?> data, String filePath, ExcelExportConfig config) throws IOException {
        ExcelExportConfig cfg = merge(config);
        try (OutputStream out = openFileOutputStream(resolvePath(filePath, cfg))) {
            export(data, out, cfg);
        }
    }

    /**
     * 单表导出到文件（使用全局默认配置）。
     *
     * @param data     数据行集合
     * @param filePath 导出文件路径
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void export(List<?> data, String filePath) throws IOException {
        export(data, filePath, null);
    }

    /**
     * 单表导出到输出流。
     *
     * @param data   数据行集合
     * @param out    目标输出流
     * @param config 导出配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    @Override
    public void export(List<?> data, OutputStream out, ExcelExportConfig config) throws IOException {
        ExcelExportConfig cfg = merge(config);
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(cfg.getSheetName());
            writeSheetContent(workbook, sheet, data, cfg);
            workbook.write(out);
        }
    }

    /**
     * 单表导出到输出流（使用全局默认配置）。
     *
     * @param data 数据行集合
     * @param out  目标输出流
     * @throws IOException 写入失败时抛出
     */
    @Override
    public void export(List<?> data, OutputStream out) throws IOException {
        export(data, out, null);
    }

    /**
     * 单表导出为字节数组。
     *
     * @param data   数据行集合
     * @param config 导出配置（可为 null）
     * @return xlsx 文件字节数组
     */
    @Override
    public byte[] exportToBytes(List<?> data, ExcelExportConfig config) {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            export(data, out, config);
            return out.toByteArray();
        } catch (IOException e) {
            throw new IllegalStateException("导出 Excel 失败", e);
        }
    }

    /**
     * 单表导出为字节数组（使用全局默认配置）。
     *
     * @param data 数据行集合
     * @return xlsx 文件字节数组
     */
    @Override
    public byte[] exportToBytes(List<?> data) {
        return exportToBytes(data, null);
    }

    // ============================================================
    // 二、多 Sheet 导出
    // ============================================================

    /**
     * 多 Sheet 导出到文件。
     *
     * @param sheets       Sheet 列表
     * @param filePath     导出文件路径
     * @param globalConfig 全局配置（可为 null）
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void exportMultiSheet(List<SheetData> sheets, String filePath, ExcelExportConfig globalConfig)
            throws IOException {
        ExcelExportConfig cfg = merge(globalConfig);
        try (OutputStream out = openFileOutputStream(resolvePath(filePath, cfg))) {
            exportMultiSheet(sheets, out, cfg);
        }
    }

    /**
     * 多 Sheet 导出到文件（使用全局默认配置）。
     *
     * @param sheets   Sheet 列表
     * @param filePath 导出文件路径
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void exportMultiSheet(List<SheetData> sheets, String filePath) throws IOException {
        exportMultiSheet(sheets, filePath, null);
    }

    /**
     * 多 Sheet 导出到输出流。
     *
     * @param sheets       Sheet 列表
     * @param out          目标输出流
     * @param globalConfig 全局配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    @Override
    public void exportMultiSheet(List<SheetData> sheets, OutputStream out, ExcelExportConfig globalConfig)
            throws IOException {
        ExcelExportConfig cfg = merge(globalConfig);
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            if (sheets != null) {
                for (SheetData sheetData : sheets) {
                    // 每个 Sheet 的独立配置优先于全局配置，其余回落至全局配置
                    ExcelExportConfig sheetCfg = merge(sheetData.getConfig(), cfg);
                    String sheetName = sheetData.getSheetName() != null && !sheetData.getSheetName().isEmpty()
                            ? sheetData.getSheetName() : sheetCfg.getSheetName();
                    Sheet sheet = workbook.createSheet(sheetName);
                    writeSheetContent(workbook, sheet, sheetData.getData(),
                            sheetData.getDataClass(), sheetCfg);
                }
            }
            workbook.write(out);
        }
    }

    /**
     * 多 Sheet 导出为字节数组。
     *
     * @param sheets       Sheet 列表
     * @param globalConfig 全局配置（可为 null）
     * @return xlsx 文件字节数组
     */
    @Override
    public byte[] exportMultiSheetToBytes(List<SheetData> sheets, ExcelExportConfig globalConfig) {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            exportMultiSheet(sheets, out, globalConfig);
            return out.toByteArray();
        } catch (IOException e) {
            throw new IllegalStateException("导出 Excel 失败", e);
        }
    }

    /**
     * 多 Sheet 导出为字节数组（使用全局默认配置）。
     *
     * @param sheets Sheet 列表
     * @return xlsx 文件字节数组
     */
    @Override
    public byte[] exportMultiSheetToBytes(List<SheetData> sheets) {
        return exportMultiSheetToBytes(sheets, null);
    }

    // ============================================================
    // 三、大数据量分批导出（SXSSF 流式写入）
    // ============================================================

    /**
     * 大数据量导出到文件（流式写入）。
     *
     * @param data     数据行集合
     * @param filePath 导出文件路径
     * @param config   导出配置（可为 null）
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void exportLarge(List<?> data, String filePath, ExcelExportConfig config) throws IOException {
        ExcelExportConfig cfg = merge(config);
        try (OutputStream out = openFileOutputStream(resolvePath(filePath, cfg))) {
            exportLarge(data, out, cfg);
        }
    }

    /**
     * 大数据量导出到输出流（流式写入）。
     *
     * @param data   数据行集合
     * @param out    目标输出流
     * @param config 导出配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    @Override
    public void exportLarge(List<?> data, OutputStream out, ExcelExportConfig config) throws IOException {
        ExcelExportConfig cfg = merge(config);
        SXSSFWorkbook workbook = new SXSSFWorkbook(cfg.getWindowRows());
        try {
            Sheet sheet = workbook.createSheet(cfg.getSheetName());
            writeSheetContent(workbook, sheet, data, cfg);
            workbook.write(out);
        } finally {
            // 释放 SXSSF 临时文件资源
            workbook.dispose();
            workbook.close();
        }
    }

    /**
     * 大数据量导出为字节数组（流式写入）。
     *
     * @param data   数据行集合
     * @param config 导出配置（可为 null）
     * @return xlsx 文件字节数组
     */
    @Override
    public byte[] exportLargeToBytes(List<?> data, ExcelExportConfig config) {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            exportLarge(data, out, config);
            return out.toByteArray();
        } catch (IOException e) {
            throw new IllegalStateException("导出 Excel 失败", e);
        }
    }

    /**
     * 分批拉取大数据量导出到文件（推荐用于数据库大数据量导出）。
     *
     * @param provider 分批数据提供者
     * @param filePath 导出文件路径
     * @param config   导出配置（可为 null）
     * @throws IOException 文件写入失败时抛出
     */
    @Override
    public void exportLargeByBatch(BatchDataProvider<?> provider, String filePath, ExcelExportConfig config)
            throws IOException {
        ExcelExportConfig cfg = merge(config);
        try (OutputStream out = openFileOutputStream(resolvePath(filePath, cfg))) {
            exportLargeByBatch(provider, out, cfg);
        }
    }

    /**
     * 分批拉取大数据量导出到输出流。
     *
     * @param provider 分批数据提供者
     * @param out      目标输出流
     * @param config   导出配置（可为 null）
     * @throws IOException 写入失败时抛出
     */
    @Override
    public void exportLargeByBatch(BatchDataProvider<?> provider, OutputStream out, ExcelExportConfig config)
            throws IOException {
        ExcelExportConfig cfg = merge(config);
        SXSSFWorkbook workbook = new SXSSFWorkbook(cfg.getWindowRows());
        try {
            Sheet sheet = workbook.createSheet(cfg.getSheetName());
            writeSheetByBatch(workbook, sheet, provider, cfg);
            workbook.write(out);
        } finally {
            workbook.dispose();
            workbook.close();
        }
    }

    // ============================================================
    // 内部实现：Sheet 内容写入
    // ============================================================

    /**
     * 将数据集合写入单个 Sheet（含标题行、表头行、数据行、列宽与冻结设置）。
     *
     * @param workbook 工作簿
     * @param sheet    目标 Sheet
     * @param data     数据行集合
     * @param cfg      合并后的导出配置
     */
    private void writeSheetContent(Workbook workbook, Sheet sheet, List<?> data, ExcelExportConfig cfg) {
        // 数据集合为空时无法推断类型，尝试使用配置中显式声明的数据类
        Class<?> dataClass = resolveDataClass(data, cfg);
        writeSheetContent(workbook, sheet, data, dataClass, cfg);
    }

    /**
     * 将数据集合写入单个 Sheet（显式指定数据类版本）。
     *
     * @param workbook  工作簿
     * @param sheet     目标 Sheet
     * @param data      数据行集合
     * @param dataClass 数据行类（可为 null）
     * @param cfg       合并后的导出配置
     */
    private void writeSheetContent(Workbook workbook, Sheet sheet, List<?> data,
                                   Class<?> dataClass, ExcelExportConfig cfg) {
        // 显式声明的数据类为空时，从数据集合首元素推断类型
        if (dataClass == null) {
            dataClass = resolveDataClass(data, cfg);
        }
        // 解析列元数据（注解驱动；无注解时回退到全部字段）
        List<ColumnMeta> columns = resolveColumns(dataClass);
        SheetWriteContext ctx = new SheetWriteContext(workbook, columns, cfg);

        int rowIndex = 0;
        // 1. 标题行（可选）：合并整行单元格，使用标题样式
        rowIndex = writeTitleRow(sheet, columns, cfg, ctx, rowIndex);
        // 2. 表头行：使用表头样式
        rowIndex = writeHeaderRow(sheet, columns, ctx, rowIndex);

        // 3. 数据行
        if (data != null) {
            for (Object rowData : data) {
                writeDataRow(sheet, columns, rowData, rowIndex++, ctx);
            }
        }

        // 4. 列宽与冻结设置
        finalizeSheet(sheet, columns, cfg, ctx, rowIndex);
    }

    /**
     * 分批拉取并写入 Sheet 内容（大数据量分批导出专用）。
     *
     * @param workbook 工作簿
     * @param sheet    目标 Sheet
     * @param provider 分批数据提供者
     * @param cfg      合并后的导出配置
     */
    private void writeSheetByBatch(Workbook workbook, Sheet sheet, BatchDataProvider<?> provider,
                                   ExcelExportConfig cfg) {
        int batchSize = cfg.getBatchSize();
        // 列元数据：优先使用配置显式声明的数据类；否则等首批数据到达后推断
        List<ColumnMeta> columns = resolveColumns(cfg.getDataClass());
        boolean columnsReady = !columns.isEmpty();
        SheetWriteContext ctx = new SheetWriteContext(workbook, columns, cfg);

        int rowIndex = 0;
        int offset = 0;
        while (true) {
            List<?> batch = provider.fetch(offset, batchSize);
            if (batch == null || batch.isEmpty()) {
                break;
            }
            // 首批数据到达后推断列结构并写入标题/表头
            if (!columnsReady) {
                Class<?> dataClass = resolveDataClass(batch, cfg);
                columns = resolveColumns(dataClass);
                ctx = new SheetWriteContext(workbook, columns, cfg);
                rowIndex = writeTitleRow(sheet, columns, cfg, ctx, rowIndex);
                rowIndex = writeHeaderRow(sheet, columns, ctx, rowIndex);
                columnsReady = !columns.isEmpty();
            }
            for (Object rowData : batch) {
                writeDataRow(sheet, columns, rowData, rowIndex++, ctx);
            }
            offset += batch.size();
        }

        finalizeSheet(sheet, columns, cfg, ctx, rowIndex);
    }

    /**
     * 写入标题行（合并首行所有列，应用标题样式）。
     *
     * @param sheet     目标 Sheet
     * @param columns   列元数据列表
     * @param cfg       合并后的导出配置
     * @param ctx       写入上下文
     * @param rowIndex  当前行下标
     * @return 下一行下标
     */
    private int writeTitleRow(Sheet sheet, List<ColumnMeta> columns, ExcelExportConfig cfg,
                              SheetWriteContext ctx, int rowIndex) {
        String title = cfg.getTitle();
        if (title == null || title.isEmpty() || columns.isEmpty()) {
            return rowIndex;
        }
        Row titleRow = sheet.createRow(rowIndex);
        titleRow.setHeightInPoints(30);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(title);
        titleCell.setCellStyle(ctx.titleStyle());
        // 标题跨所有列合并
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, columns.size() - 1));
        return rowIndex + 1;
    }

    /**
     * 写入表头行（列名行，应用表头样式）。
     *
     * @param sheet    目标 Sheet
     * @param columns  列元数据列表
     * @param ctx      写入上下文
     * @param rowIndex 当前行下标
     * @return 下一行下标
     */
    private int writeHeaderRow(Sheet sheet, List<ColumnMeta> columns, SheetWriteContext ctx, int rowIndex) {
        if (columns.isEmpty()) {
            return rowIndex;
        }
        Row headerRow = sheet.createRow(rowIndex);
        headerRow.setHeightInPoints(20);
        for (int i = 0; i < columns.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns.get(i).name());
            cell.setCellStyle(ctx.headerStyle());
            // 表头宽度参与列宽计算
            ctx.updateWidth(i, columns.get(i).name());
        }
        return rowIndex + 1;
    }

    /**
     * 写入一行数据。
     *
     * @param sheet    目标 Sheet
     * @param columns  列元数据列表
     * @param rowData  行数据对象
     * @param rowIndex 行下标
     * @param ctx      写入上下文
     */
    private void writeDataRow(Sheet sheet, List<ColumnMeta> columns, Object rowData,
                              int rowIndex, SheetWriteContext ctx) {
        if (rowData == null || columns.isEmpty()) {
            return;
        }
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(18);
        for (int i = 0; i < columns.size(); i++) {
            ColumnMeta column = columns.get(i);
            Object value = column.read(rowData);
            Cell cell = row.createCell(i);
            ctx.setCellValue(cell, value, column, i);
        }
    }

    /**
     * Sheet 收尾：设置列宽（自动或指定）与冻结表头。
     *
     * @param sheet    目标 Sheet
     * @param columns  列元数据列表
     * @param cfg      合并后的导出配置
     * @param ctx      写入上下文
     * @param rowIndex 当前行下标（用于计算冻结行数）
     */
    private void finalizeSheet(Sheet sheet, List<ColumnMeta> columns, ExcelExportConfig cfg,
                               SheetWriteContext ctx, int rowIndex) {
        if (columns.isEmpty()) {
            return;
        }
        // 列宽：优先使用注解指定宽度，否则使用自动计算值
        for (int i = 0; i < columns.size(); i++) {
            int width = columns.get(i).width() > 0 ? columns.get(i).width() : ctx.widths[i];
            sheet.setColumnWidth(i, Math.min(width + 2, 255) * 256);
        }
        // 冻结表头：有标题行时冻结 2 行（标题 + 表头），否则冻结 1 行（表头）
        if (cfg.getFreezeHeader() && rowIndex > 1) {
            int freezeRows = cfg.getTitle() != null && !cfg.getTitle().isEmpty() ? 2 : 1;
            sheet.createFreezePane(0, freezeRows);
        }
    }

    // ============================================================
    // 内部实现：列解析
    // ============================================================

    /**
     * 解析数据行类：优先取配置显式声明，其次取数据集合中首个非空元素类型。
     *
     * @param data 数据集合
     * @param cfg  合并后的导出配置
     * @return 数据行类；无法推断时返回 null
     */
    private Class<?> resolveDataClass(List<?> data, ExcelExportConfig cfg) {
        if (cfg.getDataClass() != null) {
            return cfg.getDataClass();
        }
        if (data != null) {
            for (Object item : data) {
                if (item != null) {
                    return item.getClass();
                }
            }
        }
        return null;
    }

    /**
     * 解析导出列元数据列表（已排序、已过滤隐藏列）。
     *
     * <p>规则：</p>
     * <ol>
     *   <li>遍历整个继承链上的字段（父类字段在前）；</li>
     *   <li>若存在 {@code @ExcelColumn} 标注字段，则仅使用标注字段（过滤静态/隐藏），
     *       按 {@code order} 升序排列（相同 order 保持字段声明顺序）；</li>
     *   <li>否则回退为全部非静态、非 transient、非 {@code @ExcelIgnore} 字段，列名取字段名。</li>
     * </ol>
     *
     * @param dataClass 数据行类；为 null 时返回空列表
     * @return 列元数据列表
     */
    private List<ColumnMeta> resolveColumns(Class<?> dataClass) {
        if (dataClass == null) {
            return Collections.emptyList();
        }
        // 收集继承链上的全部字段，父类在前、子类在后，保持声明顺序
        List<Class<?>> hierarchy = new ArrayList<>();
        for (Class<?> clazz = dataClass; clazz != null && clazz != Object.class; clazz = clazz.getSuperclass()) {
            hierarchy.add(clazz);
        }
        Collections.reverse(hierarchy);
        List<Field> allFields = new ArrayList<>();
        for (Class<?> clazz : hierarchy) {
            allFields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        }

        // 过滤：非静态、非 transient、非合成字段
        List<Field> candidates = allFields.stream()
                .filter(f -> !Modifier.isStatic(f.getModifiers()))
                .filter(f -> !Modifier.isTransient(f.getModifiers()))
                .filter(f -> !f.isSynthetic())
                .toList();

        // 判断是否存在注解标注字段
        List<Field> annotated = candidates.stream()
                .filter(f -> f.isAnnotationPresent(ExcelColumn.class))
                .toList();

        List<ColumnMeta> metas = new ArrayList<>();
        if (!annotated.isEmpty()) {
            // 注解驱动模式：仅导出标注字段
            for (Field field : annotated) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                if (annotation.hidden()) {
                    continue;
                }
                metas.add(new ColumnMeta(field,
                        annotation.name().isEmpty() ? field.getName() : annotation.name(),
                        annotation.order(),
                        annotation.width(),
                        annotation.dateFormat(),
                        annotation.numberFormat(),
                        annotation.align()));
            }
        } else {
            // 回退模式：导出全部非忽略字段，列名取字段名
            for (Field field : candidates) {
                if (field.isAnnotationPresent(ExcelIgnore.class)) {
                    continue;
                }
                metas.add(new ColumnMeta(field, field.getName(), 0, 0, "", "", Align.AUTO));
            }
        }

        // 按列顺序升序排列（稳定排序：相同 order 保持声明顺序）
        metas.sort(Comparator.comparingInt(ColumnMeta::order));
        return metas;
    }

    // ============================================================
    // 内部实现：配置合并与路径处理
    // ============================================================

    /**
     * 合并导出配置：将 null 字段回落至全局默认配置（{@code excel.export.*}）。
     *
     * @param config 方法级配置（可为 null）
     * @return 合并后的配置（所有字段均已解析为最终值）
     */
    private ExcelExportConfig merge(ExcelExportConfig config) {
        ExcelExportConfig.Builder builder = ExcelExportConfig.builder();
        if (config != null) {
            builder.fileName(config.getFileName() != null ? config.getFileName() : properties.getDefaultFileName());
            builder.sheetName(config.getSheetName() != null ? config.getSheetName() : properties.getDefaultSheetName());
            builder.title(config.getTitle() != null ? config.getTitle() : properties.getDefaultTitle());
            builder.freezeHeader(config.getFreezeHeader() != null ? config.getFreezeHeader() : properties.isFreezeHeader());
            builder.autoColumnWidth(config.getAutoColumnWidth() != null ? config.getAutoColumnWidth() : properties.isAutoColumnWidth());
            builder.batchSize(config.getBatchSize() != null ? config.getBatchSize() : properties.getBatchSize());
            builder.windowRows(config.getWindowRows() != null ? config.getWindowRows() : properties.getWindowRows());
            builder.booleanTrueText(config.getBooleanTrueText() != null ? config.getBooleanTrueText() : properties.getBooleanTrueText());
            builder.booleanFalseText(config.getBooleanFalseText() != null ? config.getBooleanFalseText() : properties.getBooleanFalseText());
            builder.dataClass(config.getDataClass());
            builder.styleProvider(config.getStyleProvider());
        } else {
            builder.fileName(properties.getDefaultFileName());
            builder.sheetName(properties.getDefaultSheetName());
            builder.title(properties.getDefaultTitle());
            builder.freezeHeader(properties.isFreezeHeader());
            builder.autoColumnWidth(properties.isAutoColumnWidth());
            builder.batchSize(properties.getBatchSize());
            builder.windowRows(properties.getWindowRows());
            builder.booleanTrueText(properties.getBooleanTrueText());
            builder.booleanFalseText(properties.getBooleanFalseText());
        }
        return builder.build();
    }

    /**
     * 合并导出配置（带基准配置版本）。
     *
     * <p>用于多 Sheet 导出的子配置合并：子配置中显式设置的值优先，
     * 其余字段回落至基准配置（基准配置本身已与全局默认配置合并）。</p>
     *
     * @param config 子配置（可为 null）
     * @param base   基准配置（已合并过全局默认值，不可为 null）
     * @return 合并后的配置
     */
    private ExcelExportConfig merge(ExcelExportConfig config, ExcelExportConfig base) {
        ExcelExportConfig.Builder builder = ExcelExportConfig.builder();
        builder.fileName(config != null && config.getFileName() != null ? config.getFileName() : base.getFileName());
        builder.sheetName(config != null && config.getSheetName() != null ? config.getSheetName() : base.getSheetName());
        builder.title(config != null && config.getTitle() != null ? config.getTitle() : base.getTitle());
        builder.freezeHeader(config != null && config.getFreezeHeader() != null ? config.getFreezeHeader() : base.getFreezeHeader());
        builder.autoColumnWidth(config != null && config.getAutoColumnWidth() != null ? config.getAutoColumnWidth() : base.getAutoColumnWidth());
        builder.batchSize(config != null && config.getBatchSize() != null ? config.getBatchSize() : base.getBatchSize());
        builder.windowRows(config != null && config.getWindowRows() != null ? config.getWindowRows() : base.getWindowRows());
        builder.booleanTrueText(config != null && config.getBooleanTrueText() != null ? config.getBooleanTrueText() : base.getBooleanTrueText());
        builder.booleanFalseText(config != null && config.getBooleanFalseText() != null ? config.getBooleanFalseText() : base.getBooleanFalseText());
        builder.dataClass(config != null && config.getDataClass() != null ? config.getDataClass() : base.getDataClass());
        builder.styleProvider(config != null && config.getStyleProvider() != null ? config.getStyleProvider() : base.getStyleProvider());
        return builder.build();
    }

    /**
     * 解析输出文件路径：目录自动拼接默认文件名，无扩展名自动追加 .xlsx。
     *
     * @param filePath 调用方传入的路径（可为 null）
     * @param cfg      合并后的导出配置
     * @return 最终文件路径
     */
    private String resolvePath(String filePath, ExcelExportConfig cfg) {
        String path = filePath;
        if (path == null || path.trim().isEmpty()) {
            path = cfg.getFileName();
        }
        // 以分隔符结尾视为目录，拼接默认文件名
        if (path.endsWith("/") || path.endsWith("\\")) {
            path = path + cfg.getFileName();
        }
        // 文件名部分无扩展名时追加 .xlsx
        String name = path.substring(path.lastIndexOf('/') + 1);
        if (!name.contains(".")) {
            path = path + ".xlsx";
        }
        return path;
    }

    /**
     * 打开目标文件输出流（自动创建父目录）。
     *
     * @param path 文件路径
     * @return 文件输出流
     * @throws IOException 文件不可写时抛出
     */
    private OutputStream openFileOutputStream(String path) throws IOException {
        File file = new File(path);
        File parent = file.getAbsoluteFile().getParentFile();
        if (parent != null && !parent.exists() && !parent.mkdirs()) {
            throw new IOException("创建导出目录失败: " + parent);
        }
        return new java.io.FileOutputStream(file);
    }

    // ============================================================
    // 内部类：Sheet 写入上下文（样式缓存 + 列宽累计）
    // ============================================================

    /**
     * 单次 Sheet 写入上下文。
     *
     * <p>持有工作簿相关的样式提供者、样式缓存与列宽累计数组，
     * 保证同一 Sheet 内相同（格式 + 对齐）的数据单元格复用同一个 CellStyle。</p>
     */
    private final class SheetWriteContext {

        /**
         * 当前工作簿。
         */
        private final Workbook workbook;

        /**
         * 列元数据列表（决定列数）。
         */
        private final List<ColumnMeta> columns;

        /**
         * 合并后的导出配置。
         */
        private final ExcelExportConfig cfg;

        /**
         * 样式提供者（自定义样式或内置默认样式）。
         */
        private final ExcelStyleProvider styleProvider;

        /**
         * 标题样式缓存。
         */
        private CellStyle titleStyle;

        /**
         * 表头样式缓存。
         */
        private CellStyle headerStyle;

        /**
         * 各列累计最大展示宽度（字符数）。
         */
        private final int[] widths;

        /**
         * 已参与宽度采样的数据行数（超过上限后停止采样）。
         */
        private int sampledRows;

        /**
         * 数据样式缓存：key = "对齐|格式"。
         */
        private final Map<String, CellStyle> dataStyleCache = new HashMap<>();

        /**
         * 构造写入上下文。
         *
         * @param workbook 工作簿
         * @param columns  列元数据列表
         * @param cfg      合并后的导出配置
         */
        SheetWriteContext(Workbook workbook, List<ColumnMeta> columns, ExcelExportConfig cfg) {
            this.workbook = workbook;
            this.columns = columns;
            this.cfg = cfg;
            this.styleProvider = cfg.getStyleProvider() != null ? cfg.getStyleProvider() : new DefaultStyleProvider();
            this.widths = new int[columns.size()];
        }

        /**
         * 获取标题样式（懒加载，自定义样式为空时回落到内置默认样式）。
         *
         * @return 标题样式
         */
        CellStyle titleStyle() {
            if (titleStyle == null) {
                CellStyle custom = styleProvider.createTitleStyle(workbook);
                titleStyle = custom != null ? custom : new DefaultStyleProvider().createTitleStyle(workbook);
            }
            return titleStyle;
        }

        /**
         * 获取表头样式（懒加载，自定义样式为空时回落到内置默认样式）。
         *
         * @return 表头样式
         */
        CellStyle headerStyle() {
            if (headerStyle == null) {
                CellStyle custom = styleProvider.createHeaderStyle(workbook);
                headerStyle = custom != null ? custom : new DefaultStyleProvider().createHeaderStyle(workbook);
            }
            return headerStyle;
        }

        /**
         * 更新指定列的累计最大宽度（前 {@link #WIDTH_SAMPLE_ROW_LIMIT} 行参与采样）。
         *
         * @param columnIndex 列下标
         * @param text        该单元格的展示文本
         */
        void updateWidth(int columnIndex, String text) {
            if (columnIndex >= widths.length) {
                return;
            }
            int displayWidth = displayWidth(text);
            if (displayWidth > widths[columnIndex]) {
                widths[columnIndex] = displayWidth;
            }
        }

        /**
         * 计算文本的展示宽度（中文字符按 2 个单位估算）。
         *
         * @param text 文本
         * @return 展示宽度
         */
        private int displayWidth(String text) {
            if (text == null) {
                return 0;
            }
            int width = 0;
            for (int i = 0; i < text.length(); i++) {
                char ch = text.charAt(i);
                // CJK 统一表意文字、兼容表意文字、全角符号按 2 个字符宽计算
                if ((ch >= 0x2E80 && ch <= 0x9FFF)
                        || (ch >= 0xF900 && ch <= 0xFAFF)
                        || (ch >= 0xFF00 && ch <= 0xFFEF)) {
                    width += 2;
                } else {
                    width += 1;
                }
            }
            return width;
        }

        /**
         * 向单元格写入值并应用样式（含类型转换、日期/数值格式与对齐）。
         *
         * @param cell        目标单元格
         * @param value       字段值
         * @param column      列元数据
         * @param columnIndex 列下标（用于列宽采样）
         */
        void setCellValue(Cell cell, Object value, ColumnMeta column, int columnIndex) {
            // 数据行宽度采样：仅前 N 行参与计算，避免大数据量时开销过大
            boolean sampleWidth = sampledRows < WIDTH_SAMPLE_ROW_LIMIT;
            if (value == null) {
                cell.setCellValue("");
                cell.setCellStyle(dataStyle(null, Align.AUTO));
                if (sampleWidth) {
                    updateWidth(columnIndex, "");
                }
                return;
            }

            // 对齐方式：列注解优先，AUTO 时按值类型自动映射
            Align align = resolveAlign(column, value);

            if (value instanceof Date date) {
                // 日期：按列配置的格式（默认 yyyy-MM-dd HH:mm:ss）
                String format = formatDate(column, false);
                cell.setCellValue(date);
                cell.setCellStyle(dataStyle(format, align));
                if (sampleWidth) {
                    updateWidth(columnIndex, format);
                }
            } else if (value instanceof LocalDateTime dateTime) {
                String format = formatDate(column, false);
                cell.setCellValue(Date.from(dateTime.atZone(ZoneId.systemDefault()).toInstant()));
                cell.setCellStyle(dataStyle(format, align));
                if (sampleWidth) {
                    updateWidth(columnIndex, format);
                }
            } else if (value instanceof LocalDate localDate) {
                String format = formatDate(column, true);
                cell.setCellValue(Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant()));
                cell.setCellStyle(dataStyle(format, align));
                if (sampleWidth) {
                    updateWidth(columnIndex, format);
                }
            } else if (value instanceof Boolean bool) {
                // 布尔：转换为「是 / 否」（可通过配置自定义文案）
                String text = bool ? cfg.getBooleanTrueText() : cfg.getBooleanFalseText();
                cell.setCellValue(text);
                cell.setCellStyle(dataStyle(null, align));
                if (sampleWidth) {
                    updateWidth(columnIndex, text);
                }
            } else if (value instanceof Number number) {
                // 数字：按列配置的数值格式（默认普通数值）
                cell.setCellValue(number.doubleValue());
                cell.setCellStyle(dataStyle(null, align, column.numberFormat()));
                if (sampleWidth) {
                    updateWidth(columnIndex, String.valueOf(number));
                }
            } else {
                // 其他类型：toString 输出
                String text = value.toString();
                cell.setCellValue(text);
                cell.setCellStyle(dataStyle(null, align));
                if (sampleWidth) {
                    updateWidth(columnIndex, text);
                }
            }
            sampledRows++;
        }

        /**
         * 解析单元格对齐方式：列注解指定时优先，AUTO 时按值类型映射
         * （数字右对齐、日期/布尔居中、其余左对齐）。
         *
         * @param column 列元数据
         * @param value  字段值
         * @return 对齐方式
         */
        private Align resolveAlign(ColumnMeta column, Object value) {
            if (column.align() != null && column.align() != Align.AUTO) {
                return column.align();
            }
            if (value instanceof Number) {
                return Align.RIGHT;
            }
            if (value instanceof Date || value instanceof LocalDateTime
                    || value instanceof LocalDate || value instanceof Boolean) {
                return Align.CENTER;
            }
            return Align.LEFT;
        }

        /**
         * 计算日期格式（列配置优先，无配置时按类型使用默认格式）。
         *
         * @param column  列元数据
         * @param isLocalDate 是否为纯日期类型
         * @return 日期格式字符串
         */
        private String formatDate(ColumnMeta column, boolean isLocalDate) {
            if (column.dateFormat() != null && !column.dateFormat().isEmpty()) {
                return column.dateFormat();
            }
            return isLocalDate ? DEFAULT_DATE_FORMAT : DEFAULT_DATETIME_FORMAT;
        }

        /**
         * 获取数据样式（按格式缓存；默认对齐由值类型决定）。
         *
         * @param format 日期/数值格式，可为 null
         * @param align  对齐方式（AUTO 由值类型映射）
         * @return 数据单元格样式
         */
        private CellStyle dataStyle(String format, Align align) {
            return dataStyle(format, align, null);
        }

        /**
         * 获取数据样式（按格式缓存；默认对齐由值类型决定）。
         *
         * @param format       日期/数值格式，可为 null
         * @param align        对齐方式（AUTO 由值类型映射）
         * @param numberFormat 数值格式（与 format 二选一使用，优先 format）
         * @return 数据单元格样式
         */
        private CellStyle dataStyle(String format, Align align, String numberFormat) {
            String effectiveFormat = (format != null && !format.isEmpty()) ? format : numberFormat;
            HorizontalAlignment horizontal = toHorizontalAlignment(align);
            String cacheKey = (horizontal == null ? "AUTO" : horizontal.name()) + "|"
                    + (effectiveFormat == null ? "" : effectiveFormat);
            CellStyle style = dataStyleCache.get(cacheKey);
            if (style == null) {
                // 基础数据样式：优先使用自定义提供者，否则内置默认
                CellStyle custom = styleProvider.createDataStyle(workbook);
                CellStyle base = custom != null ? custom : new DefaultStyleProvider().createDataStyle(workbook);
                style = workbook.createCellStyle();
                style.cloneStyleFrom(base);
                if (horizontal != null) {
                    style.setAlignment(horizontal);
                }
                if (effectiveFormat != null && !effectiveFormat.isEmpty()) {
                    style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(effectiveFormat));
                }
                dataStyleCache.put(cacheKey, style);
            }
            return style;
        }

        /**
         * 将对齐枚举映射为 POI 水平对齐（AUTO 返回 null，表示不覆盖基础样式对齐）。
         *
         * @param align 对齐枚举
         * @return POI 水平对齐
         */
        private HorizontalAlignment toHorizontalAlignment(Align align) {
            if (align == null || align == Align.AUTO) {
                return null;
            }
            return switch (align) {
                case LEFT -> HorizontalAlignment.LEFT;
                case CENTER -> HorizontalAlignment.CENTER;
                case RIGHT -> HorizontalAlignment.RIGHT;
                default -> null;
            };
        }
    }
}
