# Excel 导出工具 API 文档

> 组件版本：1.0.0 · 适用 Spring Boot 2.7.x / 3.x

本文档覆盖 `spring-boot-excel-tools-starter` 的全部公开 API：注解、模型、服务接口、样式 SPI、Web 便捷导出，以及完整使用示例。

---

## 目录

1. [组件结构](#1-组件结构)
2. [注解 API](#2-注解-api)
3. [导出配置模型 ExcelExportConfig](#3-导出配置模型-excelexportconfig)
4. [多 Sheet 数据模型 SheetData](#4-多-sheet-数据模型-sheetdata)
5. [统一导出服务 ExcelExportService](#5-统一导出服务-excelexportservice)
6. [分批数据提供者 BatchDataProvider](#6-分批数据提供者-batchdataprovider)
7. [自定义样式 ExcelStyleProvider](#7-自定义样式-excelstyleprovider)
8. [Web 便捷导出 ExcelWebExporter](#8-web-便捷导出-excelwebexporter)
9. [全局配置项](#9-全局配置项)
10. [完整使用示例](#10-完整使用示例)

---

## 1. 组件结构

```
io.exceltools
├── annotation                 # 注解
│   ├── ExcelColumn            # 导出列注解（列名、顺序、格式等）
│   └── ExcelIgnore            # 导出忽略注解
├── config                     # 自动配置
│   ├── ExcelExportAutoConfiguration   # 自动配置类（引入即用）
│   └── ExcelExportProperties          # 全局配置属性（excel.export.*）
├── model                      # 模型
│   ├── Align                  # 对齐方式枚举
│   ├── ExcelExportConfig      # 单次导出配置（Builder 模式）
│   └── SheetData              # 多 Sheet 导出的单个 Sheet 模型
├── service                    # 服务
│   ├── ExcelExportService     # 统一导出服务接口（核心 API）
│   ├── BatchDataProvider      # 分批数据提供者（函数式接口）
│   └── impl.DefaultExcelExportService  # 默认实现
├── style                      # 样式
│   ├── ExcelStyleProvider     # 自定义样式 SPI
│   └── DefaultStyleProvider   # 内置默认样式
└── web                        # Web 支持（可选）
    └── ExcelWebExporter       # HTTP 响应下载便捷导出器
```

---

## 2. 注解 API

### 2.1 `@ExcelColumn` — 导出列注解

标注在实体字段上，声明该字段如何导出。

| 属性 | 类型 | 默认值 | 说明 |
| ---- | ---- | ------ | ---- |
| `name` | String | 字段名 | 列名（表头显示文字） |
| `order` | int | 0 | 列顺序，值越小越靠左；相同值按字段声明顺序 |
| `width` | int | 0 | 列宽（字符数）；0 表示自动计算（按表头与数据内容估算，中文按 2 字符宽） |
| `dateFormat` | String | `""` | 日期格式（如 `yyyy-MM-dd HH:mm:ss`）；仅对 Date / LocalDateTime / LocalDate 生效，为空时用默认格式 |
| `numberFormat` | String | `""` | 数值格式（如 `#,##0.00`、`0.00%`）；仅对 Number 生效 |
| `hidden` | boolean | false | 是否隐藏该列（不导出） |
| `align` | Align | AUTO | 水平对齐方式；AUTO 时按值类型自动判断（数字右对齐、日期/布尔居中、其余左对齐） |
| `description` | String | `""` | 列说明（仅文档用途） |

**列解析规则：**

- 实体类中**存在** `@ExcelColumn` 标注字段 → 仅导出标注字段（按 `order` 排序）；
- 实体类中**不存在** `@ExcelColumn` 标注字段 → 导出全部非静态、非 transient 字段，列名取字段名；
- 两种模式下，静态字段、transient 字段均被忽略；`@ExcelIgnore` 字段不导出。

```java
public class Order {
    @ExcelColumn(name = "订单号", order = 1, width = 20)
    private String orderNo;

    @ExcelColumn(name = "金额", order = 2, numberFormat = "#,##0.00", align = Align.RIGHT)
    private BigDecimal amount;

    @ExcelColumn(name = "下单时间", order = 3, dateFormat = "yyyy-MM-dd HH:mm:ss")
    private LocalDateTime createTime;

    @ExcelColumn(name = "已支付", order = 4)
    private Boolean paid;

    @ExcelIgnore
    private String internalRemark;   // 不导出
}
```

### 2.2 `@ExcelIgnore` — 导出忽略注解

标注在字段上表示不导出。当实体类**没有任何** `@ExcelColumn` 注解（走全字段导出）时，用本注解排除个别字段。

### 2.3 `Align` — 对齐枚举

| 常量 | 说明 |
| ---- | ---- |
| `AUTO` | 自动：数字右对齐、日期/布尔居中、其余左对齐 |
| `LEFT` | 左对齐 |
| `CENTER` | 居中对齐 |
| `RIGHT` | 右对齐 |

---

## 3. 导出配置模型 `ExcelExportConfig`

通过 `ExcelExportConfig.builder()` 链式构建，控制**单次导出**行为。所有字段均可选：设置的值优先，未设置的项自动回落至 `excel.export.*` 全局配置。

| 字段 | 类型 | 对应全局配置 | 说明 |
| ---- | ---- | ------------ | ---- |
| `fileName` | String | `default-file-name` | 导出文件名（可含路径；无扩展名自动追加 `.xlsx`；仅目录时拼接默认文件名） |
| `sheetName` | String | `default-sheet-name` | Sheet 名称 |
| `title` | String | `default-title` | 标题行文字；null 表示不输出标题行 |
| `freezeHeader` | Boolean | `freeze-header` | 是否冻结表头（有标题行时冻结 2 行） |
| `autoColumnWidth` | Boolean | `auto-column-width` | 是否自动计算列宽 |
| `batchSize` | Integer | `batch-size` | 分批导出每批拉取条数 |
| `windowRows` | Integer | `window-rows` | 流式导出（SXSSF）内存窗口行数 |
| `booleanTrueText` | String | `boolean-true-text` | 布尔 true 展示文本 |
| `booleanFalseText` | String | `boolean-false-text` | 布尔 false 展示文本 |
| `dataClass` | Class<?> | — | 显式声明数据行类（数据为空时用于解析表头结构） |
| `styleProvider` | ExcelStyleProvider | — | 自定义样式提供者 |

```java
ExcelExportConfig config = ExcelExportConfig.builder()
        .sheetName("订单明细")
        .title("2026年8月订单报表")
        .freezeHeader(true)
        .autoColumnWidth(true)
        .batchSize(5000)
        .booleanTrueText("Y")
        .booleanFalseText("N")
        .build();
```

---

## 4. 多 Sheet 数据模型 `SheetData`

描述多 Sheet 导出中的**一个 Sheet**：名称、数据、可选的数据类与独立配置。

| 字段 | 类型 | 说明 |
| ---- | ---- | ---- |
| `sheetName` | String | Sheet 名称（为空时回落全局配置） |
| `data` | List<?> | 该 Sheet 的数据行集合 |
| `dataClass` | Class<?> | 数据行类（为空且数据非空时取首元素类型） |
| `config` | ExcelExportConfig | 该 Sheet 的独立配置（优先于全局配置） |

```java
List<SheetData> sheets = Arrays.asList(
        SheetData.builder().sheetName("用户").data(userList).build(),
        SheetData.builder().sheetName("订单")
                .data(orderList)
                .config(ExcelExportConfig.builder().title("订单报表").build())
                .build()
);
```

---

## 5. 统一导出服务 `ExcelExportService`

核心 API，由自动配置注册为 Spring Bean，直接注入使用。按输出目标与场景分为三类。

### 5.1 单表导出

| 方法 | 说明 |
| ---- | ---- |
| `void export(List<?> data, String filePath, ExcelExportConfig config)` | 导出到文件（config 可为 null） |
| `void export(List<?> data, String filePath)` | 导出到文件（全局默认配置） |
| `void export(List<?> data, OutputStream out, ExcelExportConfig config)` | 导出到输出流（不关闭流） |
| `void export(List<?> data, OutputStream out)` | 导出到输出流（全局默认配置） |
| `byte[] exportToBytes(List<?> data, ExcelExportConfig config)` | 导出为字节数组 |
| `byte[] exportToBytes(List<?> data)` | 导出为字节数组（全局默认配置） |

**参数说明：**

- `data`：数据行集合，元素为带 `@ExcelColumn` 注解的实体对象；
- `filePath`：目标文件路径。三种用法：
  - 完整路径：`/tmp/report/用户数据.xlsx`
  - 仅目录：`/tmp/report/`（自动拼接默认文件名）
  - 无扩展名：`/tmp/report/用户数据`（自动追加 `.xlsx`）
- `config`：单次导出配置，`null` 表示全部使用全局默认值。

```java
// 导出到文件（带标题）
exportService.export(users, "/tmp/用户数据.xlsx",
        ExcelExportConfig.builder().title("用户报表").build());

// 导出为字节数组
byte[] bytes = exportService.exportToBytes(users);
```

### 5.2 多 Sheet 导出

| 方法 | 说明 |
| ---- | ---- |
| `void exportMultiSheet(List<SheetData> sheets, String filePath, ExcelExportConfig globalConfig)` | 导出到文件 |
| `void exportMultiSheet(List<SheetData> sheets, String filePath)` | 导出到文件（全局默认配置） |
| `void exportMultiSheet(List<SheetData> sheets, OutputStream out, ExcelExportConfig globalConfig)` | 导出到输出流 |
| `byte[] exportMultiSheetToBytes(List<SheetData> sheets, ExcelExportConfig globalConfig)` | 导出为字节数组 |
| `byte[] exportMultiSheetToBytes(List<SheetData> sheets)` | 导出为字节数组（全局默认配置） |

**参数说明：**

- `sheets`：Sheet 列表，每个元素定义名称、数据与独立配置；
- `globalConfig`：全局配置，作为所有 Sheet 的公共默认值；每个 `SheetData.config` 可局部覆盖。

```java
exportService.exportMultiSheet(sheets, "/tmp/多表数据.xlsx", null);
```

### 5.3 大数据量分批导出（SXSSF 流式）

| 方法 | 说明 |
| ---- | ---- |
| `void exportLarge(List<?> data, String filePath, ExcelExportConfig config)` | 全量数据流式导出到文件 |
| `void exportLarge(List<?> data, OutputStream out, ExcelExportConfig config)` | 全量数据流式导出到输出流 |
| `byte[] exportLargeToBytes(List<?> data, ExcelExportConfig config)` | 全量数据流式导出为字节数组（注意：数组本身驻留内存，超大文件请用文件/流方式） |
| `void exportLargeByBatch(BatchDataProvider<?> provider, String filePath, ExcelExportConfig config)` | **边查边写**分批导出到文件（推荐） |
| `void exportLargeByBatch(BatchDataProvider<?> provider, OutputStream out, ExcelExportConfig config)` | 边查边写分批导出到输出流 |

**使用场景区分：**

- `exportLarge`：数据已全部在内存中（如几十万行内存集合），用 SXSSF 流式落盘，避免生成大对象；
- `exportLargeByBatch`：数据在数据库/远端，每次只拉一批到内存（如百万行），内存占用恒定，最适合数据库大表导出。

```java
// 场景一：内存全量数据
exportService.exportLarge(userList, "/tmp/all-users.xlsx",
        ExcelExportConfig.builder().windowRows(200).build());

// 场景二：数据库分批拉取（推荐，百万级数据）
BatchDataProvider<User> provider = (offset, limit) ->
        userMapper.selectPage(new Page<>(offset / limit + 1, limit),
                new LambdaQueryWrapper<>()).getRecords();

exportService.exportLargeByBatch(provider, "/tmp/db-users.xlsx",
        ExcelExportConfig.builder().batchSize(5000).build());
```

---

## 6. 分批数据提供者 `BatchDataProvider`

函数式接口，配合 `exportLargeByBatch` 使用：

```java
@FunctionalInterface
public interface BatchDataProvider<T> {
    List<T> fetch(int offset, int limit);
}
```

| 参数 | 说明 |
| ---- | ---- |
| `offset` | 当前批次起始下标（从 0 开始；第 N 批为 `(N-1) * batchSize`） |
| `limit` | 本批次最大条数（等于配置的 `batchSize`） |

**约定**：返回 `null` 或空列表表示数据已拉取完毕，导出结束。

> 提示：MyBatis-Plus 中 `offset/limit` 即分页的 `(current-1, size)` 换算关系，见上文示例。

---

## 7. 自定义样式 `ExcelStyleProvider`

样式 SPI，通过实现本接口并传入配置即可自定义三套样式：

| 方法 | 说明 |
| ---- | ---- |
| `CellStyle createTitleStyle(Workbook workbook)` | 标题样式；返回 null 用内置默认（深蓝底白字加粗居中） |
| `CellStyle createHeaderStyle(Workbook workbook)` | 表头样式；返回 null 用内置默认（中蓝底白字加粗居中带边框） |
| `CellStyle createDataStyle(Workbook workbook)` | 数据样式；返回 null 用内置默认（白底细边框垂直居中） |

三个方法都是默认方法，**可按需只覆盖其中一种**：

```java
// 只自定义表头样式，标题与数据使用默认
ExcelExportConfig config = ExcelExportConfig.builder()
        .styleProvider(new ExcelStyleProvider() {
            @Override
            public CellStyle createHeaderStyle(Workbook workbook) {
                XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                style.setFillForegroundColor(new XSSFColor(new Color(0xED, 0x7D, 0x31), null));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);
                return style;
            }
        })
        .build();
```

> 注意：数据单元格的水平对齐与日期/数值格式会在数据样式之上叠加，不影响其他属性；本组件仅导出 xlsx，工作簿实际为 `XSSFWorkbook` / `SXSSFWorkbook`，可直接使用 XSSF 专有 API（如 `XSSFColor`）。

---

## 8. Web 便捷导出 `ExcelWebExporter`

Web 环境下的下载便捷器（需 classpath 存在 Jakarta Servlet，即使用方引入了 `spring-boot-starter-web`；由自动配置按 `@ConditionalOnClass` 自动注册）。

| 方法 | 说明 |
| ---- | ---- |
| `void export(HttpServletResponse response, String fileName, List<?> data, ExcelExportConfig config)` | 单表导出并触发浏览器下载（常规数据量） |
| `void exportLarge(HttpServletResponse response, String fileName, List<?> data, ExcelExportConfig config)` | 流式导出并触发浏览器下载（大数据量） |
| `void writeResponse(HttpServletResponse response, String fileName, byte[] bytes)` | 将已有字节数组写入响应（通用） |

内部自动处理：`Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`、`Content-Disposition`（RFC 5987 UTF-8 文件名编码）、自动追加 `.xlsx` 扩展名。

```java
@GetMapping("/export/users")
public void exportUsers(HttpServletResponse response) throws IOException {
    excelWebExporter.export(response, "用户数据", userService.listAll(), null);
}
```

---

## 9. 全局配置项

前缀：`excel.export`

| 配置项 | 类型 | 默认值 | 说明 |
| ------ | ---- | ------ | ---- |
| `default-file-name` | String | `导出数据.xlsx` | 默认导出文件名 |
| `default-sheet-name` | String | `Sheet1` | 默认 Sheet 名 |
| `default-title` | String | `null` | 默认标题行（null 不输出） |
| `auto-column-width` | Boolean | `true` | 自动计算列宽 |
| `freeze-header` | Boolean | `true` | 冻结表头 |
| `batch-size` | Integer | `10000` | 分批导出每批拉取条数 |
| `window-rows` | Integer | `100` | 流式导出内存窗口行数 |
| `boolean-true-text` | String | `是` | 布尔 true 展示文本 |
| `boolean-false-text` | String | `否` | 布尔 false 展示文本 |

```yaml
excel:
  export:
    default-file-name: 导出数据.xlsx
    default-sheet-name: Sheet1
    auto-column-width: true
    freeze-header: true
    batch-size: 10000
    window-rows: 100
    boolean-true-text: 是
    boolean-false-text: 否
```

---

## 10. 完整使用示例

### 示例一：Web 接口导出单表（含标题、冻结表头）

```java
@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private ExcelWebExporter excelWebExporter;
    @Autowired
    private UserService userService;

    @GetMapping("/export")
    public void export(HttpServletResponse response) throws IOException {
        ExcelExportConfig config = ExcelExportConfig.builder()
                .sheetName("用户列表")
                .title("系统用户数据报表")
                .freezeHeader(true)
                .build();
        excelWebExporter.export(response, "用户数据", userService.listAll(), config);
    }
}
```

### 示例二：多 Sheet 报表导出到服务器

```java
public void exportMonthlyReport() throws IOException {
    List<SheetData> sheets = new ArrayList<>();
    sheets.add(SheetData.builder().sheetName("订单")
            .data(orderService.listByMonth(LocalDate.now()))
            .build());
    sheets.add(SheetData.builder().sheetName("退款")
            .data(refundService.listByMonth(LocalDate.now()))
            .build());

    exportService.exportMultiSheet(sheets, "/data/reports/月报.xlsx",
            ExcelExportConfig.builder().title("月度运营报表").build());
}
```

### 示例三：百万级订单分批导出

```java
public void exportAllOrders(String filePath) throws IOException {
    BatchDataProvider<Order> provider = (offset, limit) ->
            orderMapper.selectList(new LambdaQueryWrapper<Order>()
                    .last("LIMIT " + limit + " OFFSET " + offset));

    exportService.exportLargeByBatch(provider, filePath,
            ExcelExportConfig.builder().batchSize(10000).build());
}
```

### 示例四：自定义样式的导出

```java
public void exportWithCustomStyle() throws IOException {
    ExcelExportConfig config = ExcelExportConfig.builder()
            .title("自定义样式报表")
            .styleProvider(new ExcelStyleProvider() {
                @Override
                public CellStyle createHeaderStyle(Workbook workbook) {
                    XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
                    style.setFillForegroundColor(new XSSFColor(new Color(0xE8, 0x6B, 0x5A), null));
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style.setAlignment(HorizontalAlignment.CENTER);
                    style.setVerticalAlignment(VerticalAlignment.CENTER);
                    return style;
                }
            })
            .build();
    exportService.export(data, "/tmp/custom-style.xlsx", config);
}
```
