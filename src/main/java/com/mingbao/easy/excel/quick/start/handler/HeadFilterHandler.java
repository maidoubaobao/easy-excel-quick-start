package com.mingbao.easy.excel.quick.start.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import com.mingbao.easy.excel.quick.start.utils.ExcelUtil;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 表头筛选处理器
 *
 * @author ming
 * @since 2023/12/13
 */
@AllArgsConstructor
public class HeadFilterHandler implements SheetWriteHandler {

    /**
     * 在第几行上加筛选器，必须要大于0
     */
    private int row;

    /**
     * 要筛选的列数，必须要大于0
     */
    private int col;

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        if (row < 1 || col < 1) {
            return;
        }

        // 获取表格对象
        Sheet sheet = context.getWriteSheetHolder().getSheet();

        // 格式：A2:C2，A2代表第1列第2行，C2代表第3列第2行，意思是在第1列到第3列、第2行的单元格上加筛选器
        sheet.setAutoFilter(CellRangeAddress.valueOf("A" + row + ":" + ExcelUtil.numberToExcelLetter(col) + row));
    }
}
