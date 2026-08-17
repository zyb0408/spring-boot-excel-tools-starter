package io.exceltools.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel 自定义样式提供者接口（SPI）。
 *
 * <p>通过实现本接口并在 {@link io.exceltools.model.ExcelExportConfig#styleProvider(ExcelStyleProvider)}
 * 中传入，即可自定义导出时的标题样式、表头样式与数据样式。</p>
 *
 * <p>三个方法均提供默认实现（返回 null，表示使用组件内置默认样式），
 * 因此可按需只覆盖其中一种或两种样式，例如：</p>
 *
 * <pre>{@code
 * ExcelExportConfig config = ExcelExportConfig.builder()
 *         .styleProvider(new ExcelStyleProvider() {
 *             @Override
 *             public CellStyle createHeaderStyle(Workbook workbook) {
 *                 CellStyle style = workbook.createCellStyle();
 *                 style.setFillForegroundColor(IndexedColors.RED.getIndex());
 *                 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
 *                 style.setAlignment(HorizontalAlignment.CENTER);
 *                 return style;
 *             }
 *         })
 *         .build();
 * }</pre>
 *
 * @author exceltools
 * @since 1.0.0
 */
public interface ExcelStyleProvider {

    /**
     * 创建标题样式（合并标题行的样式）。
     *
     * <p>返回 null 时使用内置默认标题样式（深蓝底、白色加粗大字、居中）。</p>
     *
     * @param workbook 当前导出的工作簿对象
     * @return 标题单元格样式，可为 null
     */
    default CellStyle createTitleStyle(Workbook workbook) {
        return null;
    }

    /**
     * 创建表头样式（列名行的样式）。
     *
     * <p>返回 null 时使用内置默认表头样式（深蓝底、白色加粗字、居中、细边框）。</p>
     *
     * @param workbook 当前导出的工作簿对象
     * @return 表头单元格样式，可为 null
     */
    default CellStyle createHeaderStyle(Workbook workbook) {
        return null;
    }

    /**
     * 创建数据样式（数据行的基础样式）。
     *
     * <p>返回 null 时使用内置默认数据样式（细边框、垂直居中、自动换行关闭）。
     * 注意：数据单元格的水平对齐与日期/数值格式会在此基础上叠加，不会覆盖该样式的其他属性。</p>
     *
     * @param workbook 当前导出的工作簿对象
     * @return 数据单元格样式，可为 null
     */
    default CellStyle createDataStyle(Workbook workbook) {
        return null;
    }
}
