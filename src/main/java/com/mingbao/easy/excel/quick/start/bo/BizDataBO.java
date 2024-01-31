package com.mingbao.easy.excel.quick.start.bo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * 业务数据实体
 *
 * @author ming
 * @since 2023/12/8
 */
@Data
public class BizDataBO {

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("性别")
    private String gender;
}
