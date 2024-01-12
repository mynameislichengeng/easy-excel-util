package cn.lc.project.easy.excel.dto;

import java.util.List;

/**
 * @Description
 * @Date 2024/1/11 16:33
 * @Author by licheng01
 */
public class SheetOneRowItem {

    private String sheetName;

    private List<HeaderItem> headerItemList;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<HeaderItem> getHeaderItemList() {
        return headerItemList;
    }

    public void setHeaderItemList(List<HeaderItem> headerItemList) {
        this.headerItemList = headerItemList;
    }
}
