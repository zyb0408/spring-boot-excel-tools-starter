package io.exceltools.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;

/**
 * 内置默认 Excel 样式提供者。
 *
 * <p>提供开箱即用的三套样式：</p>
 * <ul>
 *   <li><strong>标题样式</strong>：深蓝底（#2F5597）、白色 16 号加粗字、居中、细边框；</li>
 *   <li><strong>表头样式</strong>：中蓝底（#4472C4）、白色 11 号加粗字、居中、细边框；</li>
 *   <li><strong>数据样式</strong>：白底、11 号常规字、细边框、垂直居中、自动换行关闭。</li>
 * </ul>
 *
 * <p>本组件仅导出 xlsx 文件，工作簿实际为 {@link org.apache.poi.xssf.usermodel.XSSFWorkbook}
 * 或 {@link org.apache.poi.xssf.streaming.SXSSFWorkbook}，故可直接使用 XSSF 专有 API 设置颜色。
 * 若需自定义样式，请实现 {@link ExcelStyleProvider} 并通过配置传入。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
public class DefaultStyleProvider implements ExcelStyleProvider {

    /**
     * 创建默认标题样式。
     *
     * @param workbook 工作簿对象（xlsx）
     * @return 标题样式
     */
    @Override
    public CellStyle createTitleStyle(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        // 背景色：深蓝
        style.setFillForegroundColor(new XSSFColor(new Color(0x2F, 0x55, 0x97), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 对齐：水平居中、垂直居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 边框：细边框，四边一致
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        // 字体：白色 16 号加粗
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("微软雅黑");
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        font.setColor(new XSSFColor(new Color(0xFF, 0xFF, 0xFF), null));
        style.setFont(font);
        return style;
    }

    /**
     * 创建默认表头样式。
     *
     * @param workbook 工作簿对象（xlsx）
     * @return 表头样式
     */
    @Override
    public CellStyle createHeaderStyle(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        // 背景色：中蓝
        style.setFillForegroundColor(new XSSFColor(new Color(0x44, 0x72, 0xC4), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 对齐：水平居中、垂直居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 边框：细边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        // 字体：白色 11 号加粗
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("微软雅黑");
        font.setBold(true);
        font.setFontHeightInPoints((short) 11);
        font.setColor(new XSSFColor(new Color(0xFF, 0xFF, 0xFF), null));
        style.setFont(font);
        return style;
    }

    /**
     * 创建默认数据样式。
     *
     * @param workbook 工作簿对象（xlsx）
     * @return 数据样式
     */
    @Override
    public CellStyle createDataStyle(Workbook workbook) {
        XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
        // 背景色：白色
        style.setFillForegroundColor(new XSSFColor(new Color(0xFF, 0xFF, 0xFF), null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 对齐：垂直居中，水平对齐由列配置动态叠加
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 边框：细边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        // 字体：黑色 11 号常规字
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        font.setColor(new XSSFColor(new Color(0x00, 0x00, 0x00), null));
        style.setFont(font);
        return style;
    }
}
