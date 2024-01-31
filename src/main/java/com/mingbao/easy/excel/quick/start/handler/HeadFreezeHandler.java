package com.mingbao.easy.excel.quick.start.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 表头冻结处理器
 *
 * @author ming
 * @since 2023/12/13
 */
@AllArgsConstructor
public class HeadFreezeHandler implements SheetWriteHandler {

    /**
     * 要冻结的行数，不需要冻结行设置成0
     */
    private int row;

    /**
     * 要冻结的列数，不需要冻结行设置成0
     */
    private int col;

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        if (row < 0 || col < 0 || (row == 0 && col == 0)) {
            return;
        }

        // 获取表格对象
        Sheet sheet = context.getWriteSheetHolder().getSheet();

        // 冻结前col列，冻结前row行
        sheet.createFreezePane(col, row);
    }
}
