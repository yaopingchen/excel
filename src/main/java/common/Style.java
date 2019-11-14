package common;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Style {
    String fontName() default "";
    short fontSize() default 0;
    HorizontalAlignment hAlignment() default HorizontalAlignment.CENTER;
    VerticalAlignment vAlignment() default VerticalAlignment.CENTER;
    IndexedColors fontColor() default IndexedColors.BLACK;
    IndexedColors backColor() default IndexedColors.WHITE;
    String dataFormat() default "";
    boolean wrapText() default true;
    boolean bold() default false;
    BorderStyle border() default BorderStyle.THIN;
}
