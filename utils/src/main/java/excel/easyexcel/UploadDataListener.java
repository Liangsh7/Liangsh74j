package excel.easyexcel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import excel.easyexcel.exception.ExcelParseException;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Liangsh
 * @description
 * @date 2022/7/27 11:51
 */
@Slf4j
public class UploadDataListener<T> extends AnalysisEventListener<T> {
    private static final int MAX_COUNT = 1000;

    /**
     * 原始数据
     */
    List<T> list = new ArrayList<>();

    /**
     * 正文起始行
     */
    private Integer headRowNumber;
    /**
     * 合并单元格
     */
    private List<CellExtra> extraMergeInfoList = new ArrayList<>();

    public UploadDataListener(Integer headRowNumber) {
        this.headRowNumber = headRowNumber;
    }

    /**
     * 每一条数据解析均调用该方法
     */
    @Override
    public void invoke(T data, AnalysisContext context) {
        if (list.size() > MAX_COUNT) {
            throw new ExcelParseException("超过最大解析数量：" + MAX_COUNT);
        }

        list.add(data);
    }

    /**
     * 解析完毕
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }

    /**
     * 获取数据
     */
    public List<T> getData() {
        return list;
    }

    /**
     * 解析合并单元格
     *
     * @param extra
     * @param context
     */

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        log.info("读取到了一条额外信息:{}", extra.getType());
        switch (extra.getType()) {
            case COMMENT:
                log.info("额外信息是批注,在rowIndex:{},columnIndex;{},内容是:{}", extra.getRowIndex(), extra.getColumnIndex(),
                        extra.getText());
                break;
            case HYPERLINK:
                if ("Sheet1!A1".equals(extra.getText())) {
                    log.info("额外信息是超链接,在rowIndex:{},columnIndex;{},内容是:{}", extra.getRowIndex(),
                            extra.getColumnIndex(), extra.getText());
                } else if ("Sheet2!A1".equals(extra.getText())) {
                    log.info(
                            "额外信息是超链接,而且覆盖了一个区间,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{},"
                                    + "内容是:{}",
                            extra.getFirstRowIndex(), extra.getFirstColumnIndex(), extra.getLastRowIndex(),
                            extra.getLastColumnIndex(), extra.getText());
                } else {
                    throw new ExcelParseException("错误的超连接");
                }
                break;
            case MERGE:
                log.info(
                        "额外信息是合并单元格,而且覆盖了一个区间,在firstRowIndex:{},firstColumnIndex;{},lastRowIndex:{},lastColumnIndex:{}",
                        extra.getFirstRowIndex(), extra.getFirstColumnIndex(), extra.getLastRowIndex(),
                        extra.getLastColumnIndex());
                if (extra.getRowIndex() >= headRowNumber) {
                    extraMergeInfoList.add(extra);
                }
                break;
            default:
        }

    }

    public List<CellExtra> getExtraMergeInfoList() {
        return extraMergeInfoList;
    }
}
