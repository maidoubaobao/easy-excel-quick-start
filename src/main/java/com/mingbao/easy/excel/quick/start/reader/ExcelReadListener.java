package com.mingbao.easy.excel.quick.start.reader;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.mingbao.easy.excel.quick.start.bo.BizDataBO;
import io.github.maidoubaobao.easy.tool.GsonUtil;
import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.util.Map;
import java.util.function.Consumer;

/**
 * Excel文件读取监听器
 *
 * @author ming
 * @since 2023/12/6
 */
@Slf4j
@AllArgsConstructor
public class ExcelReadListener implements ReadListener<BizDataBO> {

    /**
     * 表头处理器
     */
    private final Consumer<Map<Integer, ReadCellData<?>>> headConsumer;

    /**
     * 数据行处理器
     */
    private final Consumer<BizDataBO> rowConsumer;

    /**
     * 错误数据处理器
     */
    private final Consumer<Map<Integer, Cell>> errorConsumer;

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        // 读取表头
        headConsumer.accept(headMap);
    }

    @Override
    public void invoke(BizDataBO bizDataBO, AnalysisContext context) {
        // 读取行数据
        log.info("解析出来的行数据 index={} cellDTO={}", context.readRowHolder().getRowIndex(), GsonUtil.toJsonString(bizDataBO));
        rowConsumer.accept(bizDataBO);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        log.info("总数据行数 total={}", analysisContext.readRowHolder().getRowIndex());
    }

    @Override
    public void onException(Exception exception, AnalysisContext context) {
        // 封装失败数据
        log.warn("解析出错 row={} reason={}", GsonUtil.toJsonString(context.readRowHolder().getCellMap()), exception.getMessage());
        errorConsumer.accept(context.readRowHolder().getCellMap());
    }
}
