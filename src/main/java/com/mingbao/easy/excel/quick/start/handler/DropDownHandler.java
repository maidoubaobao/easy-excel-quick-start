package com.mingbao.easy.excel.quick.start.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Map;

/**
 * 单元格下拉选择处理器
 *
 * @author ming
 * @since 2023/12/15
 */
@AllArgsConstructor
public class DropDownHandler implements SheetWriteHandler {

    /**
     * 单元格下拉的起始行，0：第1行。可以直接设置成表头的行数
     */
    private int firstRow;

    /**
     * 单元格下拉的截止行，0：第1行。可以设置成表头的行数+数据的行数-1
     */
    private int lastRow;

    /**
     * 列下标和下拉数据的映射集合，列下标从0开始
     */
    private Map<Integer, String[]> colOptionMap;

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        // 如果下拉选择的起始行和截止行不规范，或者没有下拉数据，则不设置下拉选择
        if (firstRow < 0 || lastRow < firstRow || colOptionMap == null || colOptionMap.isEmpty()) {
            return;
        }

        // 获取表格对象
        Sheet sheet = context.getWriteSheetHolder().getSheet();
        // 获取下拉工具
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();

        colOptionMap.forEach((col, data) -> {
            // 下拉单元格范围，从firstRow～lastRow行，第 col 列
            CellRangeAddressList list = new CellRangeAddressList(firstRow, lastRow, col, col);

            // 设置下拉数据
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(data);

            // 创建校验器
            DataValidation validation = validationHelper.createValidation(constraint, list);
            // 如果输入的内容在下拉数据中找不到，弹出错误提示框
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);
        });
    }
}
