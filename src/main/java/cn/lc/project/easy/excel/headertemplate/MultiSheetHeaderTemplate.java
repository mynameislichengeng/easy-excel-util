package cn.lc.project.easy.excel.headertemplate;

import cn.lc.project.easy.excel.dto.HeaderItem;
import cn.lc.project.easy.excel.dto.SheetMultiRowItem;
import cn.lc.project.easy.excel.dto.SheetOneRowItem;
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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @Description
 * @Date 2024/1/11 16:29
 * @Author by licheng01
 */
public class MultiSheetHeaderTemplate {

    private Logger logger = LoggerFactory.getLogger(MultiSheetHeaderTemplate.class);

    /**
     * 多sheet-创建单行模板
     *
     * @param sheetList
     * @return
     */
    public void createSingleRowTemplate(String excelFilePath, List<SheetOneRowItem> sheetList) {

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

            //设置下拉筛选隐藏sheet
            Map<String, HeaderItem> hiddenMap = sheetList.stream()
                    .map(SheetOneRowItem::getHeaderItemList)
                    .flatMap(List::stream)
                    .collect(Collectors.toMap(
                            HeaderItem::getDropdownHiddenSheetName,
                            Function.identity(), (k1, k2) -> k1
                    ));
            hiddenMap.values().forEach(headerItem ->
                    CellUtil.createHiddenSheetAndRefersValue(workbook,
                            headerItem.getDropdownHiddenSheetName(),
                            headerItem.getDropdownFilteringValueList()));

            sheetList.forEach(r -> {
                Sheet sheet = workbook.createSheet(r.getSheetName());
                Row headerRow = sheet.createRow(0);
                r.getHeaderItemList().forEach(headerItem -> {
                    //设置单元格样式
                    CellUtil.settingCell(sheet, headerRow, redHeaderStyle, defaultHeaderStyle, headerItem);
                });
            });
            // 写入文件
            workbook.write(fileOut);
        } catch (IOException e) {
            logger.error("IOException", e);
            ExceptionUtil.throwException("创建excel表头失败");
        }
    }


    /**
     * 多sheet-创建单行模板
     *
     * @param sheetList
     * @return
     */
    public void createMultiRowTemplate(String excelFilePath, List<SheetMultiRowItem> sheetList) {
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

            //设置下拉筛选隐藏sheet
            List<HeaderItem> hiddenList = new ArrayList<>();
            for (SheetMultiRowItem sheetMultiRowItem : sheetList) {
                List<List<HeaderItem>> mk = sheetMultiRowItem.getHeaderItemList();
                for (List<HeaderItem> headerItemList : mk) {
                    hiddenList.addAll(headerItemList);
                }
            }
            hiddenList.stream()
                    .collect(Collectors.toMap(
                            HeaderItem::getHeaderName,
                            Function.identity(),
                            (k1, k2) -> k1))
                    .values()
                    .forEach(headerItem ->
                            CellUtil.createHiddenSheetAndRefersValue(workbook,
                                    headerItem.getDropdownHiddenSheetName(),
                                    headerItem.getDropdownFilteringValueList()));

            sheetList.forEach(r -> {
                Sheet sheet = workbook.createSheet(r.getSheetName());
                List<List<HeaderItem>> headerList = r.getHeaderItemList();
                for (int i = 0; i < headerList.size(); i++) {
                    Row headerRow = sheet.createRow(i);
                    List<HeaderItem> itemList = headerList.get(i);
                    itemList.forEach(headerItem -> {
                        //设置单元格样式
                        CellUtil.settingCell(sheet, headerRow, redHeaderStyle, defaultHeaderStyle, headerItem);
                    });
                }
            });
            // 写入文件
            workbook.write(fileOut);

        } catch (IOException e) {
            logger.error("IOException", e);
            ExceptionUtil.throwException("创建excel表头失败");
        }

    }
}
