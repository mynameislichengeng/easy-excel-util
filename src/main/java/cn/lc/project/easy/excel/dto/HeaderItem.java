package cn.lc.project.easy.excel.dto;

import java.util.List;

/**
 * @Description
 * @Date 2024/1/11 16:26
 * @Author by licheng01
 */
public class HeaderItem {
    /**
     * 位置
     */
    private int index;
    /**
     * 标题
     */
    private String headerName;
    /**
     * 是否必填
     */
    private boolean required;

    /**
     * 是否是下拉框
     */
    private boolean dropdownFiltering;

    /**
     * 下拉框的值
     */
    private List<String> dropdownFilteringValueList;

    /**
     * 下拉框的隐藏sheet
     */
    private String dropdownHiddenSheetName;

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public boolean isDropdownFiltering() {
        return dropdownFiltering;
    }

    public void setDropdownFiltering(boolean dropdownFiltering) {
        this.dropdownFiltering = dropdownFiltering;
    }

    public List<String> getDropdownFilteringValueList() {
        return dropdownFilteringValueList;
    }

    public void setDropdownFilteringValueList(List<String> dropdownFilteringValueList) {
        this.dropdownFilteringValueList = dropdownFilteringValueList;
    }

    public String getDropdownHiddenSheetName() {
        return dropdownHiddenSheetName;
    }

    public void setDropdownHiddenSheetName(String dropdownHiddenSheetName) {
        this.dropdownHiddenSheetName = dropdownHiddenSheetName;
    }
}
