package cn.lc.project.easy.excel.dto;

import java.util.List;

/**
 * @Description
 * @Date 2024/1/11 16:33
 * @Author by licheng01
 */
public class SheetMultiRowItem {

    private String sheetName;

    private List<List<HeaderItem>> headerItemList;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<List<HeaderItem>> getHeaderItemList() {
        return headerItemList;
    }

    public void setHeaderItemList(List<List<HeaderItem>> headerItemList) {
        this.headerItemList = headerItemList;
    }
}
