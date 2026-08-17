package io.exceltools.service.impl;

import io.exceltools.annotation.ExcelColumn;
import io.exceltools.model.Align;

import java.lang.reflect.Field;

/**
 * 导出列元数据（内部类，供导出服务使用）。
 *
 * <p>将 {@link ExcelColumn} 注解信息与反射字段绑定，缓存字段访问能力，
 * 在导出过程中按列元数据读取对象属性并写入单元格。</p>
 *
 * @author exceltools
 * @since 1.0.0
 */
final class ColumnMeta {

    /**
     * 对应的实体字段（已设置为可访问）。
     */
    private final Field field;

    /**
     * 列名（表头文字）。
     */
    private final String name;

    /**
     * 列顺序（值越小越靠左；相同值按字段声明顺序）。
     */
    private final int order;

    /**
     * 列宽（字符数），0 表示自动计算。
     */
    private final int width;

    /**
     * 日期格式（可能为空字符串）。
     */
    private final String dateFormat;

    /**
     * 数值格式（可能为空字符串）。
     */
    private final String numberFormat;

    /**
     * 对齐方式。
     */
    private final Align align;

    /**
     * 构造列元数据。
     *
     * @param field        实体字段
     * @param name         列名
     * @param order        列顺序
     * @param width        列宽
     * @param dateFormat   日期格式
     * @param numberFormat 数值格式
     * @param align        对齐方式
     */
    ColumnMeta(Field field, String name, int order, int width,
               String dateFormat, String numberFormat, Align align) {
        this.field = field;
        this.name = name;
        this.order = order;
        this.width = width;
        this.dateFormat = dateFormat;
        this.numberFormat = numberFormat;
        this.align = align;
        // 私有字段需要提升访问权限；业务实体类位于应用模块，通常可直接访问
        try {
            field.setAccessible(true);
        } catch (Exception ignore) {
            // 访问权限设置失败时，读取阶段会返回 null，不影响整体导出
        }
    }

    /**
     * 读取字段名。
     *
     * @return 字段名
     */
    String fieldName() {
        return field.getName();
    }

    /**
     * 获取列名。
     *
     * @return 列名
     */
    String name() {
        return name;
    }

    /**
     * 获取列顺序。
     *
     * @return 列顺序
     */
    int order() {
        return order;
    }

    /**
     * 获取列宽。
     *
     * @return 列宽
     */
    int width() {
        return width;
    }

    /**
     * 获取日期格式。
     *
     * @return 日期格式
     */
    String dateFormat() {
        return dateFormat;
    }

    /**
     * 获取数值格式。
     *
     * @return 数值格式
     */
    String numberFormat() {
        return numberFormat;
    }

    /**
     * 获取对齐方式。
     *
     * @return 对齐方式
     */
    Align align() {
        return align;
    }

    /**
     * 反射读取目标对象该字段的值。
     *
     * @param target 目标对象
     * @return 字段值；读取失败或字段不可访问时返回 null
     */
    Object read(Object target) {
        try {
            return field.get(target);
        } catch (IllegalAccessException e) {
            return null;
        }
    }
}
