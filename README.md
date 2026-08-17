# spring-boot-excel-tools-starter

基于 Apache POI 的 Spring Boot Starter 组件，**统一封装 Excel 导出能力**。引入 jar 包后无需额外配置即可使用，帮助团队成员保持导出行为一致。

## 特性

- **注解驱动**：通过 `@ExcelColumn` 注解声明导出列（列名、顺序、宽度、日期/数值格式、对齐等），无需手写导出逻辑
- **三种导出场景**：单表导出、多 Sheet 导出、大数据量分批导出（SXSSF 流式写入，百万级数据内存占用恒定）
- **统一导出 API**：所有成员通过 `ExcelExportService` 接口导出，行为一致
- **自定义样式**：标题样式、表头样式、数据样式均可通过 `ExcelStyleProvider` 自定义
- **多种输出目标**：文件路径、字节数组（`byte[]`）、输出流（`OutputStream`）、HTTP 响应下载
- **Spring Boot 自动配置**：`@AutoConfiguration` + `AutoConfiguration.imports`，引入即用；支持 Spring Boot 2.7.x 与 3.x

## 快速开始

### 1. 引入依赖

将本组件打包（`mvn install`）后，在使用方工程的 `pom.xml` 中加入：

```xml
<dependency>
    <groupId>io.exceltools</groupId>
    <artifactId>spring-boot-excel-tools-starter</artifactId>
    <version>1.0.0</version>
</dependency>
```

### 2. 定义导出实体（注解驱动）

```java
public class User {

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
    private String password;   // 不导出
    // getter / setter ...
}
```

### 3. 注入并导出

```java
@Service
public class UserService {

    @Autowired
    private ExcelExportService exportService;

    /** 导出到本地文件 */
    public void exportToFile() throws IOException {
        exportService.export(userMapper.listAll(), "/tmp/用户数据.xlsx",
                ExcelExportConfig.builder().title("用户信息报表").build());
    }

    /** 导出为字节数组返回前端 */
    public byte[] exportToBytes() {
        return exportService.exportToBytes(userMapper.listAll());
    }
}
```

### 4. Web 下载（可选，需引入 spring-boot-starter-web）

```java
@RestController
public class UserController {

    @Autowired
    private ExcelWebExporter excelWebExporter;

    @GetMapping("/export")
    public void export(HttpServletResponse response) throws IOException {
        excelWebExporter.export(response, "用户数据.xlsx", userMapper.listAll(), null);
    }
}
```

## 全局配置（application.yml）

```yaml
excel:
  export:
    default-file-name: 导出数据.xlsx   # 默认导出文件名（方法未指定时使用）
    default-sheet-name: Sheet1         # 默认 Sheet 名
    default-title:                     # 默认标题行（留空则不输出标题行）
    auto-column-width: true            # 是否自动计算列宽（中文按 2 字符宽估算）
    freeze-header: true                # 是否冻结表头
    batch-size: 10000                  # 分批导出每批拉取条数
    window-rows: 100                   # 流式导出内存窗口行数（SXSSF）
    boolean-true-text: 是              # 布尔 true 展示文本
    boolean-false-text: 否             # 布尔 false 展示文本
```

## 文档

- [API 文档（详细接口说明、参数说明、使用示例）](docs/API文档.md)

## 环境要求

| 依赖 | 版本 |
| ---- | ---- |
| Java | 17+ |
| Spring Boot | 2.7.x / 3.x |
| Apache POI | 5.2.5（内置依赖） |

## 构建

```bash
mvn clean install        # 打包并安装到本地仓库
mvn test                 # 运行单元测试
```

## License

Apache License 2.0
