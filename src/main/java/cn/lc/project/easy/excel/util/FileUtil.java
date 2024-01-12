package cn.lc.project.easy.excel.util;

import cn.lc.project.easy.excel.exception.ExceptionUtil;

import java.io.File;
import java.util.UUID;

/**
 * @Description
 * @Date 2024/1/11 15:49
 * @Author by licheng01
 */
public class FileUtil {
    private static final String TEMP_DIR = "D:\\";

    public static String createTempFile() {
        String filePath = TEMP_DIR + "" + UUID.randomUUID().toString().replace("-", "") + ".xlsx";
        File file = new File(filePath);
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            parentFile.mkdirs();
        }
        try {
            boolean create = file.createNewFile();
            if (!create) {
                ExceptionUtil.throwException("创建文件失败");
            }
            return filePath;
        } catch (Exception e) {
            ExceptionUtil.throwException("创建文件失败");
            return null;
        }
    }

}
