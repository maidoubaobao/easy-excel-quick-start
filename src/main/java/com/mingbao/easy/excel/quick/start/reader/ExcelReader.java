package com.mingbao.easy.excel.quick.start.reader;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.mingbao.easy.excel.quick.start.Main;
import com.mingbao.easy.excel.quick.start.bo.BizDataBO;
import com.mingbao.easy.excel.quick.start.bo.ReadResultBO;
import com.mingbao.easy.excel.quick.start.utils.ExcelUtil;
import io.github.maidoubaobao.easy.tool.GsonUtil;
import lombok.extern.slf4j.Slf4j;

import java.io.InputStream;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * excel文件读取器
 *
 * @author ming
 * @since 2023/12/15
 */
@Slf4j
public class ExcelReader {

    private ExcelReader() {

    }

    /**
     * 读取excel文件，并解析成一个结果实体
     *
     * @param fileName excel文件文件名，文件要放在resource目录下
     * @return 结果实体
     */
    public static ReadResultBO read(String fileName) {
        // 封装读取结果
        ReadResultBO readResultBO = new ReadResultBO();

        // 读取excel文件
        InputStream inputStream = Main.class.getClassLoader().getResourceAsStream(fileName);
        EasyExcel.read(inputStream, BizDataBO.class, new ExcelReadListener(
                headData -> parseHead(headData, readResultBO),
                rowData -> parseRow(rowData, readResultBO),
                errorData -> parseError(errorData, readResultBO)
        )).sheet().doRead();

        log.info("成功行数 fail={}", readResultBO.getReadList().size());
        log.info("失败行数 fail={}", readResultBO.getFailedList().size());
        return readResultBO;
    }

    /**
     * 封装表头
     *
     * @param headData   表头源数据
     * @param readResultBO 解析结果，这里用于封装表头
     */
    private static void parseHead(Map<Integer, ReadCellData<?>> headData, ReadResultBO readResultBO) {
        // 封装表头数据
        headData.forEach((index, cell) -> readResultBO.getHead().add(Collections.singletonList(cell.getStringValue())));
        log.info("封装表头 head={}", GsonUtil.toJsonString(readResultBO.getHead()));
    }

    /**
     * 封装行数据
     *
     * @param rowData    行数据映射的对象
     * @param readResultBO 解析结果，这里用于封装读取成功的行数据
     */
    private static void parseRow(BizDataBO rowData, ReadResultBO readResultBO) {
        // 封装行数据
        readResultBO.getReadList().add(rowData);
    }

    /**
     * 封装解析失败的数据
     *
     * @param errorData  解析失败的数据
     * @param readResultBO 解析结果，这里用于封装解析失败的数据
     */
    private static void parseError(Map<Integer, Cell> errorData, ReadResultBO readResultBO) {
        // 封装错误行，key是列下标，value是单元格内容
        Map<Integer, WriteCellData<Object>> errorRow = new HashMap<>();

        // 解析每一个单元格，生成错误行
        errorData.forEach((index, cell) -> {
            // 封装单元格内容
            WriteCellData<Object> errorCell = new WriteCellData<>();
            // 设置单元格样式
            errorCell.setWriteCellStyle(ExcelUtil.getContentCellStyle());
            // 设置单元格的类型和值
            errorCell.setType(((ReadCellData<?>) cell).getType());
            if (CellDataTypeEnum.NUMBER == ((ReadCellData<?>) cell).getType()) {
                errorCell.setNumberValue(((ReadCellData<?>) cell).getNumberValue());
            } else if (CellDataTypeEnum.BOOLEAN == ((ReadCellData<?>) cell).getType()) {
                errorCell.setBooleanValue(((ReadCellData<?>) cell).getBooleanValue());
            } else {
                errorCell.setStringValue(((ReadCellData<?>) cell).getStringValue());
            }

            errorRow.put(index, errorCell);
        });

        readResultBO.getFailedList().add(errorRow);
    }
}
