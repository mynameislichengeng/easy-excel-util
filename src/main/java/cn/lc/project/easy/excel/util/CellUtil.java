package cn.lc.project.easy.excel.util;

import cn.lc.project.easy.excel.dto.HeaderItem;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @Description
 * @Date 2024/1/11 15:46
 * @Author by licheng01
 */
public class CellUtil {

    public static void settingCell(Sheet sheet,
                                   Row headerRow,
                                   CellStyle redHeaderStyle,
                                   CellStyle defaultHeaderStyle,
                                   HeaderItem headerItem) {
        Cell billTypeCell = headerRow.createCell(headerItem.getIndex());
        billTypeCell.setCellValue(headerItem.getHeaderName());
        billTypeCell.setCellStyle(headerItem.isRequired() ? redHeaderStyle : defaultHeaderStyle);
        if (headerItem.isDropdownFiltering()) {
            settingDropDownValues(sheet, headerItem.getDropdownHiddenSheetName(), headerItem.getIndex());
        }
    }

    private static void settingDropDownValues(Sheet sheet, String hiddenCategoryName, int column) {
        // 获取上文名称内数据
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createFormulaListConstraint(hiddenCategoryName);
        // 设置下拉框位置
        CellRangeAddressList addressList = new CellRangeAddressList(1, 5000, column, column);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        // 处理Excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            // 数据校验
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        // 作用在目标sheet上
        sheet.addValidationData(dataValidation);
    }

    public static void createHiddenSheetAndRefersValue(Workbook workbook, String selectName, List<String> selectValue) {
        if (CollectionUtils.isEmpty(selectValue)) {
            selectValue = new ArrayList<>();
        }
        // 创建sheet，写入枚举项
        Sheet hideSheet = workbook.createSheet(selectName);
        for (int i = 0; i < selectValue.size(); i++) {
            hideSheet.createRow(i).createCell(0).setCellValue(selectValue.get(i));
        }
        // 创建名称，可被其他单元格引用
        Name category1Name = workbook.createName();
        category1Name.setNameName(selectName);
        // 设置名称引用的公式
        // 使用像'A1：B1'这样的相对值会导致在Microsoft Excel中使用工作簿时名称所指向的单元格的意外移动，
        // 通常使用绝对引用，例如'$A$1:$B$1'可以避免这种情况。
        // 参考： http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Name.html
        if (CollectionUtils.isNotEmpty(selectValue)) {
            category1Name.setRefersToFormula(selectName + "!" + "$A$1:$A$" + selectValue.size());
        }
        int index = workbook.getSheetIndex(hideSheet);
        workbook.setSheetHidden(index, true);
    }




}
