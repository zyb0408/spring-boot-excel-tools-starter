package io.exceltools.annotation;

import io.exceltools.model.Align;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 导出列注解。
 *
 * <p>标注在需要导出的实体类字段上，用于声明导出列的表头名称、顺序、宽度、格式等信息，
 * 实现「注解驱动」的 Excel 导出配置，无需手写每列的导出逻辑。</p>
 *
 * <h3>使用规则：</h3>
 * <ul>
 *   <li>若实体类中存在至少一个 {@code @ExcelColumn} 标注的字段，则<strong>只导出</strong>被标注的字段；
 *       未标注字段（包括被 {@link ExcelIgnore} 标注的字段）一律忽略。</li>
 *   <li>若实体类中没有任何 {@code @ExcelColumn} 标注的字段，则默认导出全部非静态、非 transient 字段，
 *       列名取字段名。</li>
 *   <li>列顺序按 {@link #order()} 升序排列；{@code order} 相同时按字段声明顺序排列。</li>
 * </ul>
 *
 * <h3>示例：</h3>
 * <pre>{@code
 * public class User {
 *     @ExcelColumn(name = "用户ID", order = 1)
 *     private Long id;
 *
 *     @ExcelColumn(name = "姓名", order = 2, width = 20)
 *     private String name;
 *
 *     @ExcelColumn(name = "创建时间", order = 3, dateFormat = "yyyy-MM-dd HH:mm:ss")
 *     private LocalDateTime createTime;
 *
 *     @ExcelColumn(name = "余额", order = 4, numberFormat = "#,##0.00", align = Align.RIGHT)
 *     private BigDecimal balance;
 * }
 * }</pre>
 *
 * @author exceltools
 * @since 1.0.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

    /**
     * 列名（Excel 表头显示的文字）。
     *
     * <p>为空时默认取字段名。</p>
     *
     * @return 表头列名
     */
    String name() default "";

    /**
     * 列顺序。
     *
     * <p>值越小越靠左；相同值时按字段声明顺序排列。默认 0 表示完全按声明顺序。</p>
     *
     * @return 列顺序（升序）
     */
    int order() default 0;

    /**
     * 列宽度（按字符数计算）。
     *
     * <p>取值为 0 时表示自动计算列宽：取「表头长度」与「该列数据最大展示长度」的较大值，
     * 中文按 2 个字符宽度估算，并受 Excel 单列最大宽度（255 字符）限制。</p>
     *
     * @return 列宽（字符数），0 表示自动
     */
    int width() default 0;

    /**
     * 日期格式，例如 {@code "yyyy-MM-dd HH:mm:ss"}。
     *
     * <p>仅对 {@link java.util.Date}、{@link java.time.LocalDateTime}、{@link java.time.LocalDate}
     * 类型的字段生效。为空时使用默认格式：
     * 日期时间类型为 {@code "yyyy-MM-dd HH:mm:ss"}，日期类型为 {@code "yyyy-MM-dd"}。</p>
     *
     * @return 日期格式字符串
     */
    String dateFormat() default "";

    /**
     * 数值格式，例如 {@code "#,##0.00"}、{@code "0.00%"}。
     *
     * <p>仅对数字类型字段（{@link Number} 的子类）生效。为空时以普通数值写入单元格。</p>
     *
     * @return 数值格式字符串
     */
    String numberFormat() default "";

    /**
     * 该列是否隐藏（不导出）。
     *
     * <p>隐藏列不会出现在 Excel 中，也不会参与列宽计算与表头生成。</p>
     *
     * @return true 表示隐藏该列
     */
    boolean hidden() default false;

    /**
     * 单元格对齐方式。
     *
     * <p>取值为 {@link Align#AUTO} 时按值类型自动判断：数字右对齐、其余左对齐。</p>
     *
     * @return 对齐方式
     */
    Align align() default Align.AUTO;

    /**
     * 列说明（仅供文档阅读，不影响导出结果）。
     *
     * @return 列说明
     */
    String description() default "";
}
