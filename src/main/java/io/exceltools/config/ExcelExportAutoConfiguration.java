package io.exceltools.config;

import io.exceltools.service.ExcelExportService;
import io.exceltools.service.impl.DefaultExcelExportService;
import io.exceltools.web.ExcelWebExporter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnClass;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

/**
 * Excel 导出工具自动配置类。
 *
 * <p>遵循 Spring Boot Starter 自动配置约定：使用方只需引入本 jar 包，
 * 无需任何额外配置，即自动注册 {@link ExcelExportService} Bean 并生效。</p>
 *
 * <p>注册机制：</p>
 * <ul>
 *   <li>Spring Boot 3.x：读取 {@code META-INF/spring/org.springframework.boot.autoconfigure.AutoConfiguration.imports}；</li>
 *   <li>Spring Boot 2.7.x：读取 {@code META-INF/spring.factories}。</li>
 * </ul>
 *
 * <p>注册的 Bean：</p>
 * <ul>
 *   <li>{@link ExcelExportService}：统一导出服务（核心）；</li>
 *   <li>{@link ExcelWebExporter}：Web 便捷导出器（仅当 classpath 中存在
 *       Jakarta Servlet 环境时注册，即使用方引入了 Web 相关依赖）。</li>
 * </ul>
 *
 * <p>两个 Bean 均带 {@code @ConditionalOnMissingBean}：若使用方已自行定义同类型 Bean，
 * 则自动配置让位，方便定制替换。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
// FIX:修复的 AutoConfiguration 缺少条件注解问题 - 添加 @ConditionalOnClass 确保 POI 依赖存在时才自动配置
@AutoConfiguration
@EnableConfigurationProperties(ExcelExportProperties.class)
@ConditionalOnClass({XSSFWorkbook.class, SXSSFWorkbook.class})
public class ExcelExportAutoConfiguration {

    /**
     * 注册统一导出服务 Bean。
     *
     * @param properties 全局导出配置（绑定 application.yml 的 excel.export.*）
     * @return 导出服务实例
     */
    @Bean
    @ConditionalOnMissingBean(ExcelExportService.class)
    public ExcelExportService excelExportService(ExcelExportProperties properties) {
        return new DefaultExcelExportService(properties);
    }

    /**
     * 注册 Web 便捷导出器 Bean。
     *
     * <p>仅在 classpath 存在 {@code HttpServletResponse}（Jakarta Servlet 环境）时注册，
     * 用于在 Controller 中一行代码完成「浏览器下载 Excel」。</p>
     *
     * @param exportService 统一导出服务
     * @return Web 导出器实例
     */
    @Bean
    @ConditionalOnMissingBean
    @ConditionalOnClass(jakarta.servlet.http.HttpServletResponse.class)
    public ExcelWebExporter excelWebExporter(ExcelExportService exportService) {
        return new ExcelWebExporter(exportService);
    }
}
