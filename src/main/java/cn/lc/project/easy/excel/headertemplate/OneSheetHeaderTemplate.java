package cn.lc.project.easy.excel.headertemplate;

import cn.lc.project.easy.excel.dto.HeaderItem;
import cn.lc.project.easy.excel.exception.ExceptionUtil;
import cn.lc.project.easy.excel.util.CellUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Description
 * @Date 2024/1/11 15:46
 * @Author by licheng01
 */
public class OneSheetHeaderTemplate {

    private Logger logger = LoggerFactory.getLogger(OneSheetHeaderTemplate.class);

    /**
     * 单sheet-创建单行模板
     *
     * @param sheetName
     * @param headerItems
     * @return
     */
    public void createSingleRowTemplate(String excelFilePath, String sheetName, List<HeaderItem> headerItems) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            Sheet sheet = workbook.createSheet(sheetName);
            // 创建CellStyle并设置颜色-红色
            CellStyle redHeaderStyle = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redHeaderStyle.setFont(redFont);
            redHeaderStyle.setAlignment(HorizontalAlignment.LEFT);

            // 创建CellStyle并设置颜色-默认颜色
            CellStyle defaultHeaderStyle = workbook.createCellStyle();
            defaultHeaderStyle.setAlignment(HorizontalAlignment.LEFT);

            List<HeaderItem> dropdownFilteringValueList = headerItems.stream()
                    .filter(HeaderItem::isDropdownFiltering)
                    .collect(Collectors.toList());

            dropdownFilteringValueList.forEach(headerItem ->
                    CellUtil.createHiddenSheetAndRefersValue(workbook,
                            headerItem.getDropdownHiddenSheetName(),
                            headerItem.getDropdownFilteringValueList()));

            Row headerRow = sheet.createRow(0);
            headerItems.forEach(headerItem -> {
                //设置单元格样式
                CellUtil.settingCell(sheet, headerRow, redHeaderStyle, defaultHeaderStyle, headerItem);
            });
            // 写入文件
            workbook.write(fileOut);
        } catch (IOException e) {
            logger.error("IOException", e);
            ExceptionUtil.throwException("创建excel表头失败");
        }
    }

    /**
     * 单sheet-创建多行模板
     *
     * @param sheetName
     * @param rowList
     * @return
     */
    public void createMultiRowTemplate(String excelFilePath, String sheetName, List<List<HeaderItem>> rowList) {

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {

            // 创建CellStyle并设置颜色-红色
            CellStyle redHeaderStyle = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redHeaderStyle.setFont(redFont);
            redHeaderStyle.setAlignment(HorizontalAlignment.LEFT);

            // 创建CellStyle并设置颜色-默认颜色
            CellStyle defaultHeaderStyle = workbook.createCellStyle();
            defaultHeaderStyle.setAlignment(HorizontalAlignment.LEFT);

            //下拉筛选
            rowList.forEach(r -> {
                List<HeaderItem> dropdownFilteringValueList = r.stream()
                        .filter(HeaderItem::isDropdownFiltering)
                        .collect(Collectors.toList());
                dropdownFilteringValueList.forEach(headerItem ->
                        CellUtil.createHiddenSheetAndRefersValue(workbook,
                                headerItem.getDropdownHiddenSheetName(),
                                headerItem.getDropdownFilteringValueList()));
            });

            //创建表头
            Sheet sheet = workbook.createSheet(sheetName);
            for (int i = 0; i < rowList.size(); i++) {
                Row headerRow = sheet.createRow(i);
                List<HeaderItem> itemList = rowList.get(i);
                for (HeaderItem headerItem : itemList) {
                    //设置单元格样式
                    CellUtil.settingCell(sheet, headerRow, redHeaderStyle, defaultHeaderStyle, headerItem);
                }
            }
            // 写入文件
            workbook.write(fileOut);
        } catch (IOException e) {
            logger.error("IOException", e);
            ExceptionUtil.throwException("创建excel表头失败");
        }
    }


}
