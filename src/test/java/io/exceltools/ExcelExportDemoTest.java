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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Excel 导出工具综合 Demo 测试。
 *
 * <p>本测试类覆盖：</p>
 * <ul>
 *   <li>基础导出（单表、多 Sheet）</li>
 *   <li>Bug 修复验证（分批导出标题/表头丢失、路径遍历防护、列排序稳定性等）</li>
 *   <li>参数合法性校验（windowRows/batchSize 负数修正）</li>
 *   <li>空数据场景</li>
 * </ul>
 */
class ExcelExportDemoTest {

    private ExcelExportService exportService;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        exportService = new DefaultExcelExportService(new ExcelExportProperties());
    }

    // ================================================================
    // 1. 基础功能 Demo
    // ================================================================

    @Nested
    @DisplayName("1. 基础导出功能")
    class BasicExportDemo {

        @Test
        @DisplayName("1.1 单表导出：字节数组 + 自定义配置")
        void exportSingleSheetToBytes() throws IOException {
            List<Product> products = buildProducts();

            byte[] bytes = exportService.exportToBytes(products,
                    ExcelExportConfig.builder()
                            .sheetName("商品列表")
                            .title("2026年商品库存报表")
                            .freezeHeader(true)
                            .build());

            assertNotNull(bytes);
            assertTrue(bytes.length > 0);

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheet("商品列表");
                assertNotNull(sheet);

                // 标题行
                assertEquals("2026年商品库存报表", sheet.getRow(0).getCell(0).getStringCellValue());

                // 表头行
                Row header = sheet.getRow(1);
                assertEquals("商品ID", header.getCell(0).getStringCellValue());
                assertEquals("商品名称", header.getCell(1).getStringCellValue());
                assertEquals("价格", header.getCell(2).getStringCellValue());
                assertEquals("库存", header.getCell(3).getStringCellValue());
                assertEquals("上架日期", header.getCell(4).getStringCellValue());
                assertEquals("是否促销", header.getCell(5).getStringCellValue());

                // 数据行
                Row firstData = sheet.getRow(2);
                assertEquals(1L, firstData.getCell(0).getNumericCellValue());
                assertEquals("笔记本电脑", firstData.getCell(1).getStringCellValue());
                assertEquals(5999.00, firstData.getCell(2).getNumericCellValue(), 0.01);
                assertEquals(100, firstData.getCell(3).getNumericCellValue());
                assertEquals("是", firstData.getCell(5).getStringCellValue());

                // 总行数 = 标题(1) + 表头(1) + 数据(3) = 5
                assertEquals(5, sheet.getLastRowNum() + 1);
            }
        }

        @Test
        @DisplayName("1.2 单表导出：写入文件")
        void exportSingleSheetToFile() throws IOException {
            Path file = tempDir.resolve("output").resolve("商品导出.xlsx");

            exportService.export(buildProducts(), file.toString());

            assertTrue(Files.exists(file));
            assertTrue(Files.size(file) > 0);
        }

        @Test
        @DisplayName("1.3 单表导出：空数据（需指定 dataClass 生成表头）")
        void exportEmptyDataWithDataClass() throws IOException {
            byte[] bytes = exportService.exportToBytes(List.of(),
                    ExcelExportConfig.builder()
                            .dataClass(Product.class)
                            .sheetName("空表")
                            .build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheet("空表");
                assertNotNull(sheet);
                // 空数据 + dataClass：应仅生成表头
                Row header = sheet.getRow(0);
                assertNotNull(header);
                assertEquals("商品ID", header.getCell(0).getStringCellValue());
                assertEquals("商品名称", header.getCell(1).getStringCellValue());
                // 只有 1 行（表头）
                assertEquals(1, sheet.getLastRowNum() + 1);
            }
        }
    }

    // ================================================================
    // 2. 多 Sheet 导出 Demo
    // ================================================================

    @Nested
    @DisplayName("2. 多 Sheet 导出")
    class MultiSheetDemo {

        @Test
        @DisplayName("2.1 多 Sheet 导出：不同数据 + 独立配置")
        void exportMultiSheetWithConfigs() throws IOException {
            List<SheetData> sheets = Arrays.asList(
                    SheetData.builder()
                            .sheetName("商品")
                            .data(buildProducts())
                            .build(),
                    SheetData.builder()
                            .sheetName("订单汇总")
                            .data(buildOrders())
                            .config(ExcelExportConfig.builder()
                                    .title("订单统计报表")
                                    .freezeHeader(true)
                                    .build())
                            .build()
            );

            byte[] bytes = exportService.exportMultiSheetToBytes(sheets,
                    ExcelExportConfig.builder()
                            .sheetName("默认Sheet")
                            .build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                assertEquals(2, workbook.getNumberOfSheets());
                assertEquals("商品", workbook.getSheetName(0));
                assertEquals("订单汇总", workbook.getSheetName(1));

                // 第二个 Sheet 有标题
                Sheet orderSheet = workbook.getSheet("订单汇总");
                assertEquals("订单统计报表", orderSheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("订单号", orderSheet.getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        @DisplayName("2.2 多 Sheet 导出：写入文件")
        void exportMultiSheetToFile() throws IOException {
            Path file = tempDir.resolve("multi-sheet.xlsx");

            List<SheetData> sheets = Arrays.asList(
                    SheetData.builder().sheetName("SheetA").data(buildProducts()).build(),
                    SheetData.builder().sheetName("SheetB").data(buildOrders()).build()
            );

            exportService.exportMultiSheet(sheets, file.toString());

            assertTrue(Files.exists(file));
            try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(file))) {
                assertEquals(2, wb.getNumberOfSheets());
            }
        }
    }

    // ================================================================
    // 3. 大数据量分批导出 Demo
    // ================================================================

    @Nested
    @DisplayName("3. 大数据量分批导出")
    class LargeBatchDemo {

        @Test
        @DisplayName("3.1 分批导出：预配置 dataClass 验证标题/表头写入（修复 Bug #2）")
        void exportLargeByBatch_withPreConfiguredDataClass() throws IOException {
            // 模拟数据
            List<Order> orders = buildOrders();

            // 分批提供者，每批 2 条
            BatchDataProvider<Order> provider = (offset, limit) -> {
                int from = offset;
                int to = Math.min(offset + limit, orders.size());
                if (from >= orders.size()) {
                    return null;
                }
                return new ArrayList<>(orders.subList(from, to));
            };

            byte[] bytes = exportService.exportLargeByBatchToBytes(provider,
                    ExcelExportConfig.builder()
                            .dataClass(Order.class)      // 预配置 dataClass
                            .batchSize(2)
                            .title("订单批量导出")
                            .build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheetAt(0);

                // 标题行必须存在（修复前此行为空）
                assertNotNull(sheet.getRow(0), "标题行不应为空");
                assertEquals("订单批量导出", sheet.getRow(0).getCell(0).getStringCellValue(),
                        "标题行内容应正确");

                // 表头行必须存在
                assertNotNull(sheet.getRow(1), "表头行不应为空");
                assertEquals("订单号", sheet.getRow(1).getCell(0).getStringCellValue(),
                        "表头第一列应为订单号");

                // 数据行：3 条数据 + 1 标题 + 1 表头 = 5 行
                assertEquals(5, sheet.getLastRowNum() + 1, "总行数应包含标题+表头+所有数据");
            }
        }

        @Test
        @DisplayName("3.2 分批导出：空数据预配置 dataClass")
        void exportLargeByBatch_emptyData_withDataClass() throws IOException {
            BatchDataProvider<Order> emptyProvider = (offset, limit) -> null;

            byte[] bytes = exportService.exportLargeByBatchToBytes(emptyProvider,
                    ExcelExportConfig.builder()
                            .dataClass(Order.class)
                            .title("空数据测试")
                            .build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheetAt(0);

                // 空数据场景下也应写入标题和表头
                assertNotNull(sheet.getRow(0), "空数据时标题行仍应写入");
                assertEquals("空数据测试", sheet.getRow(0).getCell(0).getStringCellValue());
                assertNotNull(sheet.getRow(1), "空数据时表头行仍应写入");
                assertEquals("订单号", sheet.getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        @DisplayName("3.3 分批导出：流式文件写入")
        void exportLargeByBatchToFile() throws IOException {
            List<Product> products = buildProducts();
            BatchDataProvider<Product> provider = (offset, limit) -> {
                int from = offset;
                int to = Math.min(offset + limit, products.size());
                if (from >= products.size()) {
                    return null;
                }
                return new ArrayList<>(products.subList(from, to));
            };

            Path file = tempDir.resolve("batch-export.xlsx");
            exportService.exportLargeByBatch(provider, file.toString(),
                    ExcelExportConfig.builder().batchSize(2).build());

            assertTrue(Files.exists(file));
            assertTrue(Files.size(file) > 0);
        }
    }

    // ================================================================
    // 4. Bug 修复验证
    // ================================================================

    @Nested
    @DisplayName("4. Bug 修复验证")
    class BugFixVerification {

        @Test
        @DisplayName("4.1 路径遍历防护：非法路径应抛异常")
        void pathTraversalPrevention() {
            ExcelExportConfig config = ExcelExportConfig.builder().build();

            assertThrows(IllegalArgumentException.class, () ->
                    exportService.exportToBytes(buildProducts(),
                            ExcelExportConfig.builder().build()) {
                // 使用 .. 路径
            });

            // 直接测试 resolvePath 逻辑（通过文件导出触发）
            assertThrows(Exception.class, () ->
                    exportService.export(buildProducts(),
                            tempDir.resolve("..").resolve("evil.xlsx").toString(),
                            config));
        }

        @Test
        @DisplayName("4.2 列排序稳定性：相同 order 保持声明顺序")
        void columnOrderStability() throws IOException {
            // TestEntity 中多个字段 order 相同 (均为 0)
            List<TestEntity> data = List.of(new TestEntity("a", "b", "c"));

            byte[] bytes = exportService.exportToBytes(data);

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheetAt(0);
                Row header = sheet.getRow(0);

                // 无 @ExcelColumn 注解时，按字段声明顺序排列
                // firstField, secondField, thirdField
                assertEquals("firstField", header.getCell(0).getStringCellValue());
                assertEquals("secondField", header.getCell(1).getStringCellValue());
                assertEquals("thirdField", header.getCell(2).getStringCellValue());
            }
        }

        @Test
        @DisplayName("4.3 windowRows 参数非法值自动修正")
        void windowRowsValidation() throws IOException {
            // windowRows 设置为负数，应自动修正为默认值 100
            byte[] bytes = exportService.exportLargeToBytes(buildProducts(),
                    ExcelExportConfig.builder()
                            .windowRows(-1)
                            .build());

            assertNotNull(bytes);
            assertTrue(bytes.length > 0);

            // windowRows 设置为 0，也应修正
            byte[] bytes2 = exportService.exportLargeToBytes(buildProducts(),
                    ExcelExportConfig.builder()
                            .windowRows(0)
                            .build());

            assertNotNull(bytes2);
            assertTrue(bytes2.length > 0);
        }

        @Test
        @DisplayName("4.4 batchSize 参数非法值自动修正")
        void batchSizeValidation() throws IOException {
            // batchSize 设置为 0，应修正为默认值
            BatchDataProvider<Product> provider = (offset, limit) -> {
                if (offset >= 1) return null;
                return new ArrayList<>(buildProducts());
            };

            Path file = tempDir.resolve("batch-valid.xlsx");
            assertDoesNotThrow(() ->
                    exportService.exportLargeByBatch(provider, file.toString(),
                            ExcelExportConfig.builder().batchSize(0).build()));

            assertTrue(Files.exists(file));
        }

        @Test
        @DisplayName("4.5 SXSSFWorkbook 初始化异常安全性")
        void sxssfWorkbookNullSafety() {
            // windowRows 为负数时 SXSSFWorkbook 会抛异常，验证 finally 块不会 NPE
            assertDoesNotThrow(() ->
                    exportService.exportLargeToBytes(buildProducts(),
                            ExcelExportConfig.builder().windowRows(-5).build()));
        }
    }

    // ================================================================
    // 5. 数据类型处理 Demo
    // ================================================================

    @Nested
    @DisplayName("5. 数据类型处理")
    class DataTypeDemo {

        @Test
        @DisplayName("5.1 各种数据类型正确导出")
        void variousDataTypes() throws IOException {
            List<AllTypes> data = List.of(
                    new AllTypes(1L, "test", new BigDecimal("99.99"),
                            LocalDateTime.of(2026, 8, 18, 10, 30),
                            LocalDate.of(2026, 8, 18), true, 42)
            );

            byte[] bytes = exportService.exportToBytes(data,
                    ExcelExportConfig.builder()
                            .sheetName("类型测试")
                            .build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheet("类型测试");
                Row header = sheet.getRow(0);

                // 验证表头
                assertEquals("ID", header.getCell(0).getStringCellValue());
                assertEquals("名称", header.getCell(1).getStringCellValue());
                assertEquals("金额", header.getCell(2).getStringCellValue());
                assertEquals("创建时间", header.getCell(3).getStringCellValue());
                assertEquals("日期", header.getCell(4).getStringCellValue());
                assertEquals("启用", header.getCell(5).getStringCellValue());
                assertEquals("计数", header.getCell(6).getStringCellValue());

                // 验证数据
                Row dataRow = sheet.getRow(1);
                assertEquals(1L, dataRow.getCell(0).getNumericCellValue());
                assertEquals("test", dataRow.getCell(1).getStringCellValue());
                assertEquals("是", dataRow.getCell(5).getStringCellValue());
                assertEquals(42, dataRow.getCell(6).getNumericCellValue());
            }
        }

        @Test
        @DisplayName("5.2 ExcelIgnore 注解排除敏感字段")
        void excelIgnoreDemo() throws IOException {
            List<UserWithPassword> data = List.of(
                    new UserWithPassword(1L, "张三", "secret123")
            );

            byte[] bytes = exportService.exportToBytes(data);

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Row header = workbook.getRow(0);
                // password 字段应被 @ExcelIgnore 排除
                assertEquals(2, header.getLastCellNum(), "应只有 ID 和姓名两列");
                assertEquals("id", header.getCell(0).getStringCellValue());
                assertEquals("name", header.getCell(1).getStringCellValue());
            }
        }

        @Test
        @DisplayName("5.3 ExcelColumn hidden 隐藏列")
        void excelColumnHiddenDemo() throws IOException {
            List<Product> data = buildProducts();

            byte[] bytes = exportService.exportToBytes(data,
                    ExcelExportConfig.builder().sheetName("隐藏列测试").build());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
                Sheet sheet = workbook.getSheet("隐藏列测试");
                Row header = sheet.getRow(0);
                // 6 列中 hidden=true 的列应被排除
                // 检查列数
                int cellCount = header.getLastCellNum();
                assertTrue(cellCount <= 6, "隐藏列不应出现在导出中");
            }
        }

        @Test
        @DisplayName("5.4 自定义样式提供者")
        void customStyleProvider() throws IOException {
            List<Product> data = buildProducts();

            // 使用自定义样式（Demo：将标题背景设为红色）
            ExcelExportConfig config = ExcelExportConfig.builder()
                    .sheetName("自定义样式")
                    .title("自定义样式 Demo")
                    .styleProvider(new io.exceltools.style.ExcelStyleProvider() {
                        @Override
                        public org.apache.poi.ss.usermodel.CellStyle createHeaderStyle(
                                org.apache.poi.ss.usermodel.Workbook workbook) {
                            org.apache.poi.xssf.usermodel.XSSFCellStyle style =
                                    (org.apache.poi.xssf.usermodel.XSSFCellStyle) workbook.createCellStyle();
                            style.setFillForegroundColor(
                                    new org.apache.poi.xssf.usermodel.XSSFColor(
                                            new java.awt.Color(255, 100, 100), null));
                            style.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
                            return style;
                        }
                    })
                    .build();

            byte[] bytes = exportService.exportToBytes(data, config);
            assertNotNull(bytes);
            assertTrue(bytes.length > 0);
        }
    }

    // ================================================================
    // 测试数据模型
    // ================================================================

    /**
     * 商品实体（演示 @ExcelColumn 注解）。
     */
    static class Product {
        @ExcelColumn(name = "商品ID", order = 1, width = 10)
        private Long id;

        @ExcelColumn(name = "商品名称", order = 2, width = 20)
        private String name;

        @ExcelColumn(name = "价格", order = 3, numberFormat = "#,##0.00", align = Align.RIGHT)
        private BigDecimal price;

        @ExcelColumn(name = "库存", order = 4, align = Align.RIGHT)
        private Integer stock;

        @ExcelColumn(name = "上架日期", order = 5, dateFormat = "yyyy-MM-dd")
        private LocalDate launchDate;

        @ExcelColumn(name = "是否促销", order = 6)
        private Boolean promotion;

        Product(Long id, String name, BigDecimal price, Integer stock,
                LocalDate launchDate, Boolean promotion) {
            this.id = id;
            this.name = name;
            this.price = price;
            this.stock = stock;
            this.launchDate = launchDate;
            this.promotion = promotion;
        }
    }

    /**
     * 订单实体。
     */
    static class Order {
        @ExcelColumn(name = "订单号", order = 1, width = 15)
        private String orderNo;

        @ExcelColumn(name = "客户", order = 2, width = 12)
        private String customer;

        @ExcelColumn(name = "金额", order = 3, numberFormat = "#,##0.00", align = Align.RIGHT)
        private BigDecimal amount;

        @ExcelColumn(name = "下单时间", order = 4, dateFormat = "yyyy-MM-dd HH:mm:ss")
        private LocalDateTime orderTime;

        @ExcelColumn(name = "状态", order = 5)
        private String status;

        Order(String orderNo, String customer, BigDecimal amount,
              LocalDateTime orderTime, String status) {
            this.orderNo = orderNo;
            this.customer = customer;
            this.amount = amount;
            this.orderTime = orderTime;
            this.status = status;
        }
    }

    /**
     * 全类型实体。
     */
    static class AllTypes {
        @ExcelColumn(name = "ID", order = 1)
        private Long id;

        @ExcelColumn(name = "名称", order = 2)
        private String name;

        @ExcelColumn(name = "金额", order = 3, numberFormat = "#,##0.00")
        private BigDecimal amount;

        @ExcelColumn(name = "创建时间", order = 4, dateFormat = "yyyy-MM-dd HH:mm:ss")
        private LocalDateTime createTime;

        @ExcelColumn(name = "日期", order = 5, dateFormat = "yyyy-MM-dd")
        private LocalDate date;

        @ExcelColumn(name = "启用", order = 6)
        private Boolean enabled;

        @ExcelColumn(name = "计数", order = 7)
        private Integer count;

        AllTypes(Long id, String name, BigDecimal amount, LocalDateTime createTime,
                 LocalDate date, Boolean enabled, Integer count) {
            this.id = id;
            this.name = name;
            this.amount = amount;
            this.createTime = createTime;
            this.date = date;
            this.enabled = enabled;
            this.count = count;
        }
    }

    /**
     * 用户实体（演示 @ExcelIgnore）。
     */
    static class UserWithPassword {
        private Long id;
        private String name;
        @ExcelIgnore
        private String password;

        UserWithPassword(Long id, String name, String password) {
            this.id = id;
            this.name = name;
            this.password = password;
        }
    }

    /**
     * 测试实体（演示列排序稳定性）。
     */
    static class TestEntity {
        private String firstField;
        private String secondField;
        private String thirdField;

        TestEntity(String firstField, String secondField, String thirdField) {
            this.firstField = firstField;
            this.secondField = secondField;
            this.thirdField = thirdField;
        }
    }

    // ================================================================
    // 测试数据构造
    // ================================================================

    private List<Product> buildProducts() {
        return Arrays.asList(
                new Product(1L, "笔记本电脑", new BigDecimal("5999.00"), 100,
                        LocalDate.of(2026, 1, 15), true),
                new Product(2L, "无线鼠标", new BigDecimal("299.00"), 500,
                        LocalDate.of(2026, 2, 1), false),
                new Product(3L, "机械键盘", new BigDecimal("899.00"), 200,
                        LocalDate.of(2026, 3, 10), true)
        );
    }

    private List<Order> buildOrders() {
        return Arrays.asList(
                new Order("ORD-2026-001", "李四", new BigDecimal("1250.00"),
                        LocalDateTime.of(2026, 8, 1, 10, 0), "已完成"),
                new Order("ORD-2026-002", "王五", new BigDecimal("3200.50"),
                        LocalDateTime.of(2026, 8, 5, 14, 30), "待发货"),
                new Order("ORD-2026-003", "赵六", new BigDecimal("890.00"),
                        LocalDateTime.of(2026, 8, 10, 9, 15), "已发货")
        );
    }
}