package com.mingbao.easy.excel.quick.start.writer;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import com.mingbao.easy.excel.quick.start.Main;
import com.mingbao.easy.excel.quick.start.handler.ColumnHideHandler;
import com.mingbao.easy.excel.quick.start.handler.DropDownHandler;
import com.mingbao.easy.excel.quick.start.handler.HeadFilterHandler;
import com.mingbao.easy.excel.quick.start.handler.HeadFreezeHandler;
import lombok.extern.slf4j.Slf4j;

import java.net.URL;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * excel文件输出器
 *
 * @author peter.bao
 * @since 2024/1/31
 */
@Slf4j
public class ExcelWriter {

    private ExcelWriter() {

    }

    /**
     * 输出文件
     *
     * @param head 表头
     * @param data 数据
     */
    public static void write(List<List<String>> head, List<Map<Integer, WriteCellData<Object>>> data) {
        // 输出文件路径
        String path = Optional.ofNullable(Main.class.getClassLoader().getResource("write.xlsx")).map(URL::getPath).orElse(null);
        if (path == null) {
            log.info("输出的文件不存在");
            return;
        }

        EasyExcel.write(path).head(head).sheet()
                // 设置单元格宽度：20
                .registerWriteHandler(new SimpleColumnWidthStyleStrategy(20))
                // 设置表头高度：25 单元格高度：15
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short) 15, (short) 20))
                // 隐藏列
                .registerWriteHandler(new ColumnHideHandler(null))
                // 单元格下拉选择
                .registerWriteHandler(new DropDownHandler(1, 1, Collections.singletonMap(0, new String[]{"小明", "小蓝"})))
                // 表头筛选
                .registerWriteHandler(new HeadFilterHandler(1, 1))
                // 表头冻结
                .registerWriteHandler(new HeadFreezeHandler(1, 0))
                .doWrite(data);
        log.info("输出文件 path={}", path);
    }
}
