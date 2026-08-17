package io.exceltools.model;

/**
 * 单元格水平对齐方式枚举。
 *
 * <p>用于 {@code @ExcelColumn(align = ...)} 与 {@link ExcelExportConfig} 中自定义单元格对齐。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
public enum Align {

    /**
     * 自动判断：数字类型右对齐，其余类型左对齐（推荐默认值）。
     */
    AUTO,

    /**
     * 左对齐。
     */
    LEFT,

    /**
     * 居中对齐。
     */
    CENTER,

    /**
     * 右对齐。
     */
    RIGHT
}
