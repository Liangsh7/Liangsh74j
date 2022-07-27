package excel.easyexcel;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.util.CollectionUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

/**
 * @Description: 解析excel
 * @Author: Liangsh
 * @Date: 2022/7/12 23:13
 */
@Slf4j
public class ExcelAnalysisHelper<T> {

    public List<T> getList(MultipartFile file, Class<T> clazz) {
        return getList(file, clazz, 0, 1);
    }

    /**
     * 调用该方法处理excel文件
     * 该方法一次性将所有数据读入内存后处理，需要控制最大数据量
     *
     * @param file          excel文件
     * @param clazz         解析的模板对象
     * @param sheetNo       excel工作表序号
     * @param headRowNumber 表头数
     * @return 处理后的结果集
     */
    public List<T> getList(MultipartFile file, Class<T> clazz, Integer sheetNo, Integer headRowNumber) {
        UploadDataListener<T> listener = new UploadDataListener<>(headRowNumber);
        try {
            EasyExcel.read(file.getInputStream(), clazz, listener).extraRead(CellExtraTypeEnum.MERGE).sheet(sheetNo).headRowNumber(headRowNumber).doRead();
        } catch (IOException e) {
            log.error("解析错误：" + e.getMessage());
        }
        // 获取额外单元格信息,额外单元格信息为异步处理结果
        List<CellExtra> extraMergeInfoList = listener.getExtraMergeInfoList();
        if (CollectionUtils.isEmpty(extraMergeInfoList)) {
            return listener.getData();
        }
        return parseMergeData(listener.getData(), extraMergeInfoList, headRowNumber);
    }

    /**
     * 处理合并单元格
     *
     * @param data               解析数据
     * @param extraMergeInfoList 合并单元格信息
     * @param headRowNumber      起始行
     * @return 填充好的解析数据
     */
    private List<T> parseMergeData(List<T> data, List<CellExtra> extraMergeInfoList, Integer headRowNumber) {
        // 循环所有合并单元格信息
        extraMergeInfoList.forEach(cellExtra -> {
            int firstRowIndex = cellExtra.getFirstRowIndex() - headRowNumber;
            int lastRowIndex = cellExtra.getLastRowIndex() - headRowNumber;
            int firstColumnIndex = cellExtra.getFirstColumnIndex();
            int lastColumnIndex = cellExtra.getLastColumnIndex();
            // 获取初始值
            Object initValue = getInitValueFromList(firstRowIndex, firstColumnIndex, data);
            // 设置值
            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                    setInitValueToList(initValue, i, j, data);
                }
            }
        });
        return data;
    }

    /**
     * 设置合并单元格的值
     *
     * @param filedValue  值
     * @param rowIndex    行
     * @param columnIndex 列
     * @param data        解析数据
     */
    public void setInitValueToList(Object filedValue, Integer rowIndex, Integer columnIndex, List<T> data) {
        T object = data.get(rowIndex);

        for (Field field : object.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == columnIndex) {
                    try {
                        field.set(object, filedValue);
                        break;
                    } catch (IllegalAccessException e) {
                        throw new RuntimeException("解析数据时发生异常!");
                    }
                }
            }
        }
    }


    /**
     * 获取合并单元格的初始值
     * rowIndex对应list的索引
     * columnIndex对应实体内的字段
     *
     * @param firstRowIndex    起始行
     * @param firstColumnIndex 起始列
     * @param data             列数据
     * @return 初始值
     */
    private Object getInitValueFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> data) {
        Object filedValue = null;
        T object = data.get(firstRowIndex);
        for (Field field : object.getClass().getDeclaredFields()) {
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == firstColumnIndex) {
                    try {
                        filedValue = field.get(object);
                        break;
                    } catch (IllegalAccessException e) {
                        throw new RuntimeException("解析数据时发生异常!");
                    }
                }
            }
        }
        return filedValue;
    }
}
