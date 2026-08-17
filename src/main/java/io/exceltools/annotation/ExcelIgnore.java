package io.exceltools.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 导出忽略注解。
 *
 * <p>标注在字段上表示该字段不参与 Excel 导出（例如密码、内部标识等敏感或无关字段）。</p>
 *
 * <h3>生效场景：</h3>
 * <ul>
 *   <li>实体类中<strong>没有任何</strong> {@link ExcelColumn} 注解时，默认导出全部字段，
 *       此时可通过本注解显式排除某些字段；</li>
 *   <li>实体类中<strong>存在</strong> {@link ExcelColumn} 注解时，未标注字段本身就被忽略，
 *       本注解无需使用。</li>
 * </ul>
 *
 * @author exceltools
 * @since 1.0.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelIgnore {
}
