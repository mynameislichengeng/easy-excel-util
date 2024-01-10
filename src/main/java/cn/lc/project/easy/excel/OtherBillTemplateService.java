package cn.lc.project.easy.excel;


import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.stream.Collectors;

/**
 * @author wangYan
 */

public class OtherBillTemplateService {

    private Logger logger = LoggerFactory.getLogger(OtherBillTemplateService.class);

    private static final String TEMP_DIR = "D:\\";
    private static final String SHEET_NAME = "Sheet1";
    private static final String TEMPLATE_NAME = "其他收支单据批量导入模板.xlsx";

    public static void main(String[] args) {
        OtherBillTemplateService s = new OtherBillTemplateService();
        s.downloadTemplateByInput();
        String out = s.getTemplateByOutput(new HashMap<>());
        System.out.println(out);
    }

    public void downloadTemplateByInput() {
        String filePath = createHeaderTemplateFileByInput();

    }


    public static class HeaderItem {
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

        public boolean isRequired() {
            return required;
        }

        public void setRequired(boolean required) {
            this.required = required;
        }

        public String getDropdownHiddenSheetName() {
            return dropdownHiddenSheetName;
        }

        public void setDropdownHiddenSheetName(String dropdownHiddenSheetName) {
            this.dropdownHiddenSheetName = dropdownHiddenSheetName;
        }
    }

    private String createHeaderTemplateFileByInput(List<HeaderItem> headerItems) {
        //配置，决定哪些字段作为必填

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(SHEET_NAME);
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
                    .filter(HeaderItem::isDropdownFiltering).collect(Collectors.toList());
            dropdownFilteringValueList.forEach(headerItem ->
                    createHiddenSheetAndRefersValue(workbook,
                            headerItem.getDropdownHiddenSheetName(),
                            headerItem.getDropdownFilteringValueList()));

            Row headerRow = sheet.createRow(0);
            headerItems.forEach(headerItem -> {
                //设置单元格样式
                settingCell(sheet, headerRow, redHeaderStyle, defaultHeaderStyle, headerItem);
            });
            // 保存文件
            String outPath = getTempFile();
            try (FileOutputStream fileOut = new FileOutputStream(outPath)) {
                workbook.write(fileOut);
            }
            return outPath;
        } catch (IOException e) {
            logger.error("IOException", e);
            throw new ClientException("创建模板异常");
        }
    }

    public String getTemplateByOutput(Map<String, String> configMap) {
        return createHeaderTemplateFileByOutput(configMap);
    }

    private String createHeaderTemplateFileByOutput(Map<String, String> configMap) {
        //配置，决定哪些字段作为必填
        List<String> billTypeNameList = new ArrayList<>();
        List<String> categoryNameList = new ArrayList<>();
        //缴费点
        List<String> paySiteNameList = new ArrayList<>();
        //支付方式
        List<String> payTypeNameList = new ArrayList<>();
        //对账状态
        List<String> checkStatusList = new ArrayList<>();
        //对账人、创建人
        List<String> userList = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook()) {
            // 创建CellStyle并设置颜色-红色
            CellStyle redHeaderStyle = workbook.createCellStyle();
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            redHeaderStyle.setFont(redFont);
            redHeaderStyle.setAlignment(HorizontalAlignment.LEFT);

            // 创建CellStyle并设置颜色-默认颜色
            CellStyle defaultHeaderStyle = workbook.createCellStyle();
            defaultHeaderStyle.setAlignment(HorizontalAlignment.LEFT);
            createHiddenSheetAndRefersValue(workbook, "billTypeNameList", billTypeNameList);
            createHiddenSheetAndRefersValue(workbook, "categoryNameList", categoryNameList);
            createHiddenSheetAndRefersValue(workbook, "paySiteNameList", paySiteNameList);
            createHiddenSheetAndRefersValue(workbook, "payTypeNameList", payTypeNameList);
            createHiddenSheetAndRefersValue(workbook, "checkStatusList", checkStatusList);
            createHiddenSheetAndRefersValue(workbook, "userList", userList);

            for (int i = 0; i < 2; i++) {
                Sheet sheet = workbook.createSheet(i == 0 ? "导入失败" : "导入成功");
                //设置单元格样式
                settingCell(sheet, redHeaderStyle, defaultHeaderStyle, configMap,
                        billTypeNameList, categoryNameList, paySiteNameList,
                        payTypeNameList, checkStatusList, userList, false, i == 0, true);
            }
            // 保存文件
            String outPath = getTempFile();
            try (FileOutputStream fileOut = new FileOutputStream(outPath)) {
                workbook.write(fileOut);
            }
            return outPath;
        } catch (IOException e) {
            logger.error("IOException", e);
            throw new ClientException("创建模板异常");
        }
    }

    private void settingCell(Sheet sheet,
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


    private String getTempFile() {
        String filePath = TEMP_DIR + "" + UUID.randomUUID().toString().replace("-", "") + ".xlsx";
        File file = new File(filePath);
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            parentFile.mkdirs();
        }
        try {
            boolean create = file.createNewFile();
            if (!create) {
                throw new ClientException("创建临时文件失败");
            }
            return filePath;
        } catch (Exception e) {
            logger.error("getTempFile", e);
            throw new ClientException("创建临时文件失败");
        }
    }


    private void settingDropDownValues(Sheet sheet, String hiddenCategoryName, int column) {
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

    private void createHiddenSheetAndRefersValue(Workbook workbook, String selectName, List<String> selectValue) {
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
