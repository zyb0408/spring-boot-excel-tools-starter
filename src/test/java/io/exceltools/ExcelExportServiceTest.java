package io.exceltools;

import io.exceltools.annotation.ExcelColumn;
import io.exceltools.annotation.ExcelIgnore;
import io.exceltools.config.ExcelExportProperties;
import io.exceltools.model.Align;
import io.exceltools.model.ExcelExportConfig;
import io.exceltools.model.SheetData;
import io.exceltools.service.BatchDataProvider;
import io.exceltools.service.ExcelExportService;
import io.exceltools.service.impl.DefaultExcelExportService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 默认导出服务单元测试。
 *
 * <p>覆盖三种导出场景（单表、多 Sheet、分批导出）的核心行为校验：
 * 表头、数据行、Sheet 数量、文件名解析等。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
class ExcelExportServiceTest {

    /**
     * 被测导出服务（直接实例化，脱离 Spring 容器，使用内置默认配置）。
     */
    private ExcelExportService exportService;

    /**
     * 测试前准备：初始化导出服务。
     */
    @BeforeEach
    void setUp() {
        exportService = new DefaultExcelExportService(new ExcelExportProperties());
    }

    /**
     * 测试：单表导出为字节数组，校验表头与数据内容。
     */
    @Test
    @DisplayName("单表导出：校验表头、数据行与单元格格式")
    void exportSingleSheet() throws IOException {
        byte[] bytes = exportService.exportToBytes(buildUsers(), ExcelExportConfig.builder()
                .sheetName("用户列表")
                .title("用户信息表")
                .build());

        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertEquals(1, workbook.getNumberOfSheets(), "应只生成 1 个 Sheet");
            Sheet sheet = workbook.getSheet("用户列表");
            assertTrue(sheet != null, "Sheet 名应为「用户列表」");

            // 第 0 行为标题行（合并单元格）
            assertEquals("用户信息表", sheet.getRow(0).getCell(0).getStringCellValue(), "标题行内容不正确");

            // 第 1 行为表头行
            Row header = sheet.getRow(1);
            assertEquals("用户ID", header.getCell(0).getStringCellValue());
            assertEquals("姓名", header.getCell(1).getStringCellValue());
            assertEquals("年龄", header.getCell(2).getStringCellValue());
            assertEquals("余额", header.getCell(3).getStringCellValue());
            assertEquals("注册时间", header.getCell(4).getStringCellValue());
            assertEquals("是否启用", header.getCell(5).getStringCellValue());

            // 第 2 行为第一条数据
            Row data = sheet.getRow(2);
            assertEquals(1L, data.getCell(0).getNumericCellValue());
            assertEquals("张三", data.getCell(1).getStringCellValue());
            assertEquals(25, data.getCell(2).getNumericCellValue());
            assertEquals(1000.5d, data.getCell(3).getNumericCellValue());
            assertEquals("是", data.getCell(5).getStringCellValue());

            // 共 3 条数据 + 标题行 + 表头行 = 5 行
            assertEquals(5, sheet.getLastRowNum() + 1, "总行数不正确");
        }
    }

    /**
     * 测试：多 Sheet 导出，校验 Sheet 数量与各自数据。
     */
    @Test
    @DisplayName("多 Sheet 导出：一次生成多个 Sheet")
    void exportMultiSheet() throws IOException {
        List<SheetData> sheets = Arrays.asList(
                SheetData.builder().sheetName("用户").data(buildUsers()).build(),
                SheetData.builder().sheetName("备份")
                        .data(Arrays.asList(buildUsers().get(0)))
                        .config(ExcelExportConfig.builder().title("备份表").build())
                        .build()
        );

        byte[] bytes = exportService.exportMultiSheetToBytes(sheets);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertEquals(2, workbook.getNumberOfSheets(), "应生成 2 个 Sheet");
            assertEquals("用户", workbook.getSheetName(0));
            assertEquals("备份", workbook.getSheetName(1));
            // 第一个 Sheet：无标题行，1 行表头 + 3 行数据
            assertEquals(4, workbook.getSheet("用户").getLastRowNum() + 1, "用户 Sheet 行数不正确");
            // 第二个 Sheet：1 行标题 + 1 行表头 + 1 行数据
            assertEquals(3, workbook.getSheet("备份").getLastRowNum() + 1, "备份 Sheet 行数不正确");
            assertEquals("备份表", workbook.getSheet("备份").getRow(0).getCell(0).getStringCellValue());
        }
    }

    /**
     * 测试：分批拉取导出，校验分页逻辑与总行数。
     */
    @Test
    @DisplayName("分批导出：按页拉取并完整写入")
    void exportLargeByBatch() throws IOException {
        List<?> users = buildUsers();
        // 每批 2 条，共 3 批（2 + 2 + 1）
        BatchDataProvider<Object> provider = (offset, limit) -> {
            int from = offset;
            int to = Math.min(offset + limit, users.size());
            if (from >= users.size()) {
                return null;
            }
            return new ArrayList<>(users.subList(from, to));
        };

        Path file = Files.createTempFile("batch-export", ".xlsx");
        try {
            exportService.exportLargeByBatch(provider, file.toString(),
                    ExcelExportConfig.builder().batchSize(2).build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(file))) {
                Sheet sheet = workbook.getSheetAt(0);
                // 1 行表头 + 3 行数据
                assertEquals(4, sheet.getLastRowNum() + 1, "分批导出总行数不正确");
                assertEquals("王五", sheet.getRow(3).getCell(1).getStringCellValue(), "最后一行数据不正确");
            }
        } finally {
            Files.deleteIfExists(file);
        }
    }

    /**
     * 测试：流式大数据量导出（SXSSF）。
     */
    @Test
    @DisplayName("流式导出：SXSSF 写入校验")
    void exportLarge() throws IOException {
        byte[] bytes = exportService.exportLargeToBytes(buildUsers(),
                ExcelExportConfig.builder().windowRows(2).build());
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertEquals(1, workbook.getNumberOfSheets());
            assertEquals(4, workbook.getSheetAt(0).getLastRowNum() + 1);
        }
    }

    /**
     * 测试：文件名解析（目录自动拼接默认文件名）。
     */
    @Test
    @DisplayName("文件导出：目录路径自动拼接默认文件名")
    void exportToDirectory() throws IOException {
        Path dir = Files.createTempDirectory("excel-export-test");
        try {
            // 以目录结尾的路径 → 自动拼接默认文件名「导出数据.xlsx」
            exportService.export(buildUsers(), dir.toString() + "/");
            assertTrue(Files.exists(dir.resolve("导出数据.xlsx")), "目录下应生成默认文件名");
        } finally {
            // 清理测试产物
            Files.deleteIfExists(dir.resolve("导出数据.xlsx"));
            Files.deleteIfExists(dir);
        }
    }

    /**
     * 构造测试用户数据。
     *
     * @return 用户列表
     */
    private List<User> buildUsers() {
        return Arrays.asList(
                new User(1L, "张三", 25, new BigDecimal("1000.50"),
                        LocalDateTime.of(2026, 1, 1, 10, 30), true, "pwd1"),
                new User(2L, "李四", 30, new BigDecimal("2000.00"),
                        LocalDateTime.of(2026, 2, 1, 9, 0), false, "pwd2"),
                new User(3L, "王五", 28, new BigDecimal("350.75"),
                        LocalDateTime.of(2026, 3, 15, 16, 45), true, "pwd3")
        );
    }

    /**
     * 测试用用户实体（演示 @ExcelColumn 注解用法）。
     */
    static class User {

        @ExcelColumn(name = "用户ID", order = 1, width = 10)
        private Long id;

        @ExcelColumn(name = "姓名", order = 2, width = 16)
        private String name;

        @ExcelColumn(name = "年龄", order = 3, align = Align.CENTER)
        private Integer age;

        @ExcelColumn(name = "余额", order = 4, numberFormat = "#,##0.00", align = Align.RIGHT)
        private BigDecimal balance;

        @ExcelColumn(name = "注册时间", order = 5, dateFormat = "yyyy-MM-dd HH:mm:ss")
        private LocalDateTime createTime;

        @ExcelColumn(name = "是否启用", order = 6)
        private Boolean enabled;

        @ExcelIgnore
        private String password;

        User(Long id, String name, Integer age, BigDecimal balance,
             LocalDateTime createTime, Boolean enabled, String password) {
            this.id = id;
            this.name = name;
            this.age = age;
            this.balance = balance;
            this.createTime = createTime;
            this.enabled = enabled;
            this.password = password;
        }
    }
}
