package com.mingbao.easy.excel.quick.start;

import com.mingbao.easy.excel.quick.start.bo.ReadResultBO;
import com.mingbao.easy.excel.quick.start.reader.ExcelReader;
import com.mingbao.easy.excel.quick.start.writer.ExcelWriter;
import io.github.maidoubaobao.easy.tool.GsonUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;

/**
 * @author ming
 * @since 2023/12/12
 */
@Slf4j
public class Main {

    public static void main(String[] args) {
        // 读Excel
        ReadResultBO readResultBO = ExcelReader.read("read.xlsx");
        log.info("解析成功的数据 success={}", GsonUtil.toJsonString(readResultBO.getReadList()));
        if (CollectionUtils.isNotEmpty(readResultBO.getFailedList())) {
            // 写Excel
            ExcelWriter.write(readResultBO.getHead(), readResultBO.getFailedList());
        }
    }
}
