package cn.zeroeden.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Map;

/**
 * @author Zero
 * @Description 自定义引擎解析
 */
public class ExcelExportEngine {

    /**
     * 模板引擎解析填充数据
     *
     * @param object    具体数据
     * @param workbook  被填充Excel
     * @param imagePath 如果有图片填充时的路径
     * @return
     * @throws Exception
     */
    public static Workbook writeToExcel(Object object, Workbook workbook, String imagePath) throws Exception {
        // 1. 把对象转为map
        Map<String, Object> map = EntityUtils.entityToMap(object);
        Sheet sheet = workbook.getSheetAt(0);
        // 默认遍历100行，100列 --可以根据需求改动
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < 100; i++) {
            row = sheet.getRow(i);
            if (row == null) {
                // 当前行为空
                break;
            } else {
                for (int j = 0; j < 100; j++) {
                    cell = row.getCell(j);
                    if (cell != null) {
                        // 根据具体情况填充值
                        writeCell(cell, map);
                    }
                }
            }
        }
        // 当需要导出图片时
        if (StringUtils.isNoneBlank(imagePath)) {
            // 填充图片
            // 3.1 先创建一个字节输出流
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            // 3.2 读取图片，放入了一个带有缓存区的图片类中
            BufferedImage bufferedImage = ImageIO.read(new File(imagePath));
            // 计算文件后缀名
            String extName = imagePath.substring(imagePath.lastIndexOf(".") + 1).toUpperCase();
            // 3.3 b把图片写入到字节输出流中
            ImageIO.write(bufferedImage, extName, byteArrayOutputStream);
            // 3.4 Patriarch 控制图片的写入 / ClientAnchor 指定图片的位置
            Drawing drawingPatriarch = sheet.createDrawingPatriarch();
            // 获取图片的存放位置
            Sheet sheet1 = workbook.getSheetAt(1);
            int col1 = ((Double) sheet1.getRow(0).getCell(0).getNumericCellValue()).intValue();
            int row1 = ((Double) sheet1.getRow(0).getCell(1).getNumericCellValue()).intValue();
            int col2 = ((Double) sheet1.getRow(0).getCell(2).getNumericCellValue()).intValue();
            int row2 = ((Double) sheet1.getRow(0).getCell(3).getNumericCellValue()).intValue();
            // 左上角 偏移x  偏移y  右下角 偏移x 偏移 y  后面就是表格的位置
            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col2, row2);
            // 3.5 开始把图片写入到sheet指定的位置
            int format = 0;
            switch (extName) {
                case "JPG": {
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                }
                case "JPEG": {
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                }
                case "PNG": {
                    format = XSSFWorkbook.PICTURE_TYPE_PNG;
                }
            }
            drawingPatriarch.createPicture(anchor, workbook.addPicture(byteArrayOutputStream.toByteArray(), format));

        }
        return workbook;
    }

    /**
     * 比较单元格中的值，是否和map中的key（字段值)一致,如果一致说明此单元格需要填充对应的值
     *
     * @param cell
     * @param map
     */
    private static void writeCell(Cell cell, Map<String, Object> map) {
        CellType type = cell.getCellType();
        // 判断 此单元格数据的类型--如果是公式的话不理它
        switch (type) {
            case FORMULA: {
                break;
            }
            default: {
                String value = cell.getStringCellValue();
                if (map.containsKey(value)) {
                    cell.setCellValue(map.get(value).toString());
                }
            }
        }

    }
}
