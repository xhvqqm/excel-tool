package xin.qixia.domain;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.tomcat.util.bcel.classfile.ArrayElementValue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author qixia
 */
public class ExcelBo {
    /**
     * 工作表名称
     */
    private String sheetName = "Sheet1";

    /**
     * 最终数据
     */
    private Map<String, Object> data;

    /**
     * 键值数据
     */
    private Map<String, String> map;

    /**
     * 单元格合并
     */
    List<CellRangeAddress> cellList;

    public String getSheetName() {
        return sheetName;
    }

    public ExcelBo setSheetName(String sheetName) {
        this.sheetName = sheetName;

        return this;
    }

    public Map<String, Object> getData() {
        return data;
    }

    public void setData(Map<String, Object> data) {
        this.data = data;
    }

    public Map<String, String> getMap() {
        return map;
    }

    public void setMap(Map<String, String> map) {
        this.map = map;
    }

    public List<CellRangeAddress> getCellList() {
        return cellList;
    }

    public void setCellList(List<CellRangeAddress> cellList) {
        this.cellList = cellList;
    }

    /**
     * @param key
     * @param value
     */
    public ExcelBo put(String key, String value) {
        if (map == null) map = new HashMap<>();
        if (data == null) data = new HashMap<>();
        data.computeIfAbsent("map", k -> map);
        map.put(key, value);

        return this;
    }

    /**
     * @param key
     * @param value
     */
    public <T> ExcelBo put(String key, List<T> value) {
        if (data == null) data = new HashMap<>();
        data.put(key, value);

        return this;
    }

    /**
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public void add(int firstRow, int lastRow, int firstCol, int lastCol) {
        if (cellList == null) cellList = new ArrayList<>();
        cellList.add(new CellRangeAddress(firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1));
    }
}
