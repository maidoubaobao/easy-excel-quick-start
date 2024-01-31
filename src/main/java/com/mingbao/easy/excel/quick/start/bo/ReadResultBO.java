package com.mingbao.easy.excel.quick.start.bo;

import com.alibaba.excel.metadata.data.WriteCellData;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * excel读取数据实体
 *
 * @author ming
 * @since 2023/12/15
 */
@Data
public class ReadResultBO {

    /**
     * 读取到的对象集合
     */
    private final List<BizDataBO> readList = new ArrayList<>();

    /**
     * 读取到的表头，用于生成失败数据的文件
     */
    private final List<List<String>> head = new ArrayList<>();

    /**
     * 读取失败的行集合，带表头行
     */
    private final List<Map<Integer, WriteCellData<Object>>> failedList = new ArrayList<>();
}
