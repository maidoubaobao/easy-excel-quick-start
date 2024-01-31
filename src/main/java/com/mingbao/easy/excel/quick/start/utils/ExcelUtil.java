package com.mingbao.easy.excel.quick.start.utils;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel工具类
 *
 * @author ming
 * @since 2023/12/13
 */
public class ExcelUtil {

    private ExcelUtil() {

    }

    /**
     * 将数字转成大写字母，1代表A...26代表Z，27代表AA
     */
    public static String numberToExcelLetter(int number) {
        number--;
        if (number < 26) {
            return String.valueOf((char) (number + 65));
        } else {
            int i = number / 26;
            int j = number % 26 + 1;
            return numberToExcelLetter(i) + numberToExcelLetter(j);
        }
    }

    /**
     * 生成内容样式
     */
    public static WriteCellStyle getContentCellStyle() {
        // 设置内容样式
        WriteCellStyle contentStyle = new WriteCellStyle();

        // 设置边框
        contentStyle.setBorderTop(BorderStyle.THIN);
        contentStyle.setBorderBottom(BorderStyle.THIN);
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);

        // 设置水平居中
        contentStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);

        // 设置字体大小
        WriteFont font = new WriteFont();
        font.setFontHeightInPoints((short) 12);
        contentStyle.setWriteFont(font);

        return contentStyle;
    }
}
