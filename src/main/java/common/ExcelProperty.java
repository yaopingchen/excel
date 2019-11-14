package common;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelProperty {
    int index();

    /**
     * 横向合并分组
     * @return
     */
    String groupRow() default "";

    /**
     * 纵向合并依据
     * @return
     */
    String groupBy() default "";

}
