package cn.lc.project.easy.excel.exception;

/**
 * @Description
 * @Date 2024/1/11 15:52
 * @Author by licheng01
 */
public class ExceptionUtil {


    public static void throwException(String message) {
        throw new RuntimeException(message);
    }

}
