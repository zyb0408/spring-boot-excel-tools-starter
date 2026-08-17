package io.exceltools.service;

import java.util.List;

/**
 * 分批数据提供者（函数式接口）。
 *
 * <p>用于大数据量分批导出：由调用方实现本接口，按 offset/limit 分页从数据库
 * 或其他数据源拉取数据，导出服务边拉取边写入，避免全量数据驻留内存。</p>
 *
 * <h3>使用示例（MyBatis-Plus / MyBatis 分页）：</h3>
 * <pre>{@code
 * BatchDataProvider<User> provider = (offset, limit) ->
 *         userMapper.selectPage(new Page<>(offset / limit + 1, limit),
 *                 new LambdaQueryWrapper<>()).getRecords();
 *
 * exportService.exportLargeByBatch(provider, "用户数据.xlsx",
 *         ExcelExportConfig.builder().batchSize(5000).build());
 * }</pre>
 *
 * <p>约定：返回 {@code null} 或空列表表示数据已拉取完毕，导出结束。</p>
 *
 * @param <T> 数据行类型
 * @author exceltools
 * @since 1.0.0
 */
@FunctionalInterface
public interface BatchDataProvider<T> {

    /**
     * 分页拉取一批数据。
     *
     * <p>导出服务会依次以 offset = 0, batchSize, 2*batchSize, ... 调用本方法，
     * 直至返回 {@code null} 或空列表。</p>
     *
     * @param offset 当前批次起始下标（从 0 开始，为已写入行数，即 offset/batchSize 表示页号）
     * @param limit  本批次最大条数（即配置的 batchSize）
     * @return 本批次数据列表；null 或空列表表示没有更多数据
     */
    List<T> fetch(int offset, int limit);
}
