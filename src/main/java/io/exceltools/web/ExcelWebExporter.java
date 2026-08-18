package io.exceltools.web;

import io.exceltools.model.ExcelExportConfig;
import io.exceltools.service.ExcelExportService;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;

/**
 * Web 便捷导出器。
 *
 * <p>在 Web 项目中封装「浏览器下载 Excel」的标准流程：
 * 设置响应头（Content-Type、Content-Disposition）、编码文件名、写出字节流。
 * Controller 中一行代码即可完成导出下载。</p>
 *
 * <h3>使用示例：</h3>
 * <pre>{@code
 * @RestController
 * public class UserController {
 *
 *     @Autowired
 *     private ExcelWebExporter excelWebExporter;
 *
 *     @GetMapping("/export")
 *     public void exportUser(HttpServletResponse response) throws IOException {
 *         List<User> users = userService.listAll();
 *         excelWebExporter.export(response, "用户数据.xlsx", users, null);
 *     }
 * }
 * }</pre>
 *
 * <p>注意：本类仅在 classpath 存在 Jakarta Servlet 环境时由自动配置注册，
 * 使用方需引入 {@code spring-boot-starter-web}（或等价依赖）。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
public class ExcelWebExporter {

    /**
     * xlsx 文件的 MIME 类型。
     */
    private static final String XLSX_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /**
     * 统一导出服务（委托执行实际导出逻辑）。
     */
    private final ExcelExportService exportService;

    /**
     * 构造 Web 导出器。
     *
     * @param exportService 统一导出服务
     */
    public ExcelWebExporter(ExcelExportService exportService) {
        this.exportService = exportService;
    }

    /**
     * 导出单表数据到 HTTP 响应（浏览器下载）。
     *
     * @param response HTTP 响应对象
     * @param fileName 下载文件名（自动追加 .xlsx 扩展名，浏览器端无需中文乱码处理）
     * @param data     数据行集合
     * @param config   导出配置（可为 null，使用全局默认配置）
     * @throws IOException 写出响应失败时抛出
     */
    public void export(HttpServletResponse response, String fileName, List<?> data, ExcelExportConfig config)
            throws IOException {
        byte[] bytes = exportService.exportToBytes(data, config);
        writeResponse(response, fileName, bytes);
    }

    /**
     * 导出单表数据到 HTTP 响应（流式版，适用于大数据量）。
     *
     * <p>使用 SXSSF 流式写入，内存占用与数据量无关，数据通过响应流边写边发。</p>
     *
     * @param response HTTP 响应对象
     * @param fileName 下载文件名
     * @param data     数据行集合
     * @param config   导出配置（可为 null；{@code windowRows} 控制内存窗口）
     * @throws IOException 写出响应失败时抛出
     */
    public void exportLarge(HttpServletResponse response, String fileName, List<?> data, ExcelExportConfig config)
            throws IOException {
        prepareResponse(response, fileName);
        try (OutputStream out = response.getOutputStream()) {
            exportService.exportLarge(data, out, config);
        }
    }

    /**
     * 将字节数组写入 HTTP 响应（通用方法）。
     *
     * @param response HTTP 响应对象
     * @param fileName 下载文件名
     * @param bytes    xlsx 文件字节数组
     * @throws IOException 写出响应失败时抛出
     */
    public void writeResponse(HttpServletResponse response, String fileName, byte[] bytes) throws IOException {
        prepareResponse(response, fileName);
        try (OutputStream out = response.getOutputStream()) {
            out.write(bytes);
            out.flush();
        }
    }

    /**
     * 设置响应头：Content-Type、Content-Disposition（文件名 UTF-8 编码）。
     *
     * @param response HTTP 响应对象
     * @param fileName 下载文件名
     */
    private void prepareResponse(HttpServletResponse response, String fileName) {
        String name = fileName;
        if (name == null || name.isEmpty()) {
            name = "export";
        }
        // 无扩展名时自动追加 .xlsx
        if (!name.contains(".")) {
            name = name + ".xlsx";
        }
        response.setContentType(XLSX_CONTENT_TYPE);
        response.setCharacterEncoding(StandardCharsets.UTF_8.name());
        // RFC 5987 格式：避免中文文件名乱码，兼容现代浏览器
        // URLEncoder.encode(String, Charset) 在 Java 17 中抛出受检异常，
        // 由于 StandardCharsets.UTF_8 为 JDK 内置常量，此处异常永远不会发生
        String encoded;
        try {
            encoded = URLEncoder.encode(name, StandardCharsets.UTF_8).replace("+", "%20");
        } catch (java.io.UnsupportedEncodingException e) {
            // UTF-8 为 JDK 必备字符集，此分支永不触发
            encoded = name;
        }
        response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encoded);
    }
}
