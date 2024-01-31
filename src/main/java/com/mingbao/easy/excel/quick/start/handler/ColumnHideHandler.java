package com.mingbao.easy.excel.quick.start.handler;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import lombok.AllArgsConstructor;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * 列隐藏处理器
 *
 * @author ming
 * @since 2023/12/13
 */
@AllArgsConstructor
public class ColumnHideHandler implements SheetWriteHandler {

    /**
     * 隐藏下标
     */
    private List<Integer> hideIndex;

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        if (CollectionUtils.isEmpty(hideIndex)) {
            return;
        }

        // 获取表格对象
        Sheet sheet = context.getWriteSheetHolder().getSheet();

        // 隐藏单元格
        hideIndex.forEach(index -> sheet.setColumnHidden(index, true));
    }
}
