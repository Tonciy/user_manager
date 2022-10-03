package cn.zeroeden.service;

import cn.zeroeden.mapper.UserMapper;
import cn.zeroeden.utils.ExcelExportEngine;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import cn.zeroeden.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;


import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class UserService {

    @Autowired
    private UserMapper userMapper;

    // 处理日期转化
    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page, pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        // 获取输出流
        ServletOutputStream outputStream = response.getOutputStream();
        // 创建工作簿
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
        // 创建工作表
        WritableSheet sheet = workbook.createSheet("用户表", 0);
        // 调整列宽
        sheet.setColumnView(0, 5);
        sheet.setColumnView(1, 8);
        sheet.setColumnView(2, 15);
        sheet.setColumnView(3, 15);
        sheet.setColumnView(4, 30);
        // 封装标题
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        // 填充标题
        Label label = null;
        for (int i = 0; i < titles.length; i++) {
            // 列角标、行角标、单元格内容
            label = new Label(i, 0, titles[i]);
            sheet.addCell(label);
        }

        // 查找用户数据
        List<User> userList = userMapper.selectAll();
        // 用户数据填充到工作表中
        int rowIndex = 1;
        for (User user : userList) {
            for (int i = 0; i < titles.length; i++) {
                String context = "";
                switch (i) {
                    case 0:
                        context = user.getId().toString();
                        break;
                    case 1:
                        context = user.getUserName();
                        break;
                    case 2:
                        context = user.getPhone();
                        break;
                    case 3:
                        context = simpleDateFormat.format(user.getHireDate());
                        break;
                    case 4:
                        context = user.getAddress();
                        break;
                    default:
                        context = "";
                }
                label = new Label(i, rowIndex, context);
                sheet.addCell(label);
            }
            rowIndex++;
        }
        // 文件导出 一个流两个头(文件类型，文件的打开方式)
        response.setContentType("application/vnd.ms-excel");
        String filename = "jxl入门-用户表.xls";
        // 指定编码时因为浏览器对中文支持不好，而ISO8859-1浏览器就可以识别的了
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 导出
        workbook.write();
        workbook.close();
        outputStream.close();
    }

    public void uploadExcel(MultipartFile file) throws Exception {
        // 获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        // 获取对应工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 获取数据最后一行角标
        int lastRowIndex = sheet.getLastRowNum();
        // 获取数据
        User user = null;
        XSSFRow row = null;
        for (int i = 1; i <= lastRowIndex; i++) {
            row = sheet.getRow(i);
            String userName = row.getCell(0).getStringCellValue();
            String phone = row.getCell(1).getStringCellValue();
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            Integer salary = ((Double) (row.getCell(4).getNumericCellValue())).intValue();
            Date hireDate = simpleDateFormat.parse(row.getCell(5).getStringCellValue());
            Date birthday = simpleDateFormat.parse(row.getCell(6).getStringCellValue());
            String address = row.getCell(7).getStringCellValue();
            user = new User();
            user.setUserName(userName);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthday);
            user.setAddress(address);
            System.out.println(user);
            userMapper.insert(user);
        }
    }

    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        // 创建对应的工作簿即工作表
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("用户数据");
        // 设置列框  1为一个标准字母的 256分之一
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        // 填充标题行
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = null;
        for (int i = 0; i < titles.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(titles[i]);
        }
        // 填充具体数据
        List<User> users = userMapper.selectAll();
        int rowIndex = 1;
        XSSFRow dataRow = null;
        for (User user : users) {
            // 创建这一行
            dataRow = sheet.createRow(rowIndex);
            // 填充这一行
            dataRow.createCell(0).setCellValue(user.getId());
            dataRow.createCell(1).setCellValue(user.getUserName());
            dataRow.createCell(2).setCellValue(user.getPhone());
            dataRow.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
            dataRow.createCell(4).setCellValue(user.getAddress());
            rowIndex++;
        }
        // 设置文件打开方式
        String filename = "用户表-POI-Excel导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 设置文件类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // 导出
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

    }

    public void downLoadXlsxByPoiWithStyle(HttpServletResponse response) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("带样式的用户表");
        XSSFRow titleRow = sheet.createRow(0);
        // 设置标题行高
        titleRow.setHeightInPoints(42);
        // 设置列宽
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        // 需求：1、边框线 全边框 2、行高 42  3、合并单元格  第一行的前5个  4、对齐方式：水平垂直都要居中
        // 设置单元格样式
        XSSFCellStyle titleRowCellStyle = workbook.createCellStyle();
        titleRowCellStyle.setBorderBottom(BorderStyle.THIN);
        titleRowCellStyle.setBorderLeft(BorderStyle.THIN);
        titleRowCellStyle.setBorderRight(BorderStyle.THIN);
        titleRowCellStyle.setBorderTop(BorderStyle.THIN);
        // 设置对齐方式
        titleRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置字体样式
        XSSFFont font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);
        titleRowCellStyle.setFont(font);
        for (int i = 0; i < 5; i++) {
            XSSFCell cell = titleRow.createCell(i);
            cell.setCellStyle(titleRowCellStyle);
        }
        // 合并单元格  firstRow  endRow   firstColumn   endColumn
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        // 填充标题行数据
//        titleRow.createCell(0).setCellValue("用户信息数据");
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据");

        // 设置小标题样式
        XSSFCellStyle smallRowCellStyle = workbook.createCellStyle();
        // 1. 克隆已存在的样式--减少代码
        smallRowCellStyle.cloneStyleFrom(titleRowCellStyle);
//        smallRowCellStyle.setBorderBottom(BorderStyle.THIN);
//        smallRowCellStyle.setBorderLeft(BorderStyle.THIN);
//        smallRowCellStyle.setBorderRight(BorderStyle.THIN);
//        smallRowCellStyle.setBorderTop(BorderStyle.THIN);
//        // 设置对齐方式
//        smallRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
//        smallRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 2. 设置小标题的字体
        XSSFFont smallFront = new XSSFFont();
        smallFront.setFontName("黑体");
        smallFront.setFontHeight((short) 12);
        smallFront.setBold(true);
        smallRowCellStyle.setFont(smallFront);


        // 填充小标题行
        String[] titles = {"编号", "姓名", "手机号", "入职日期", "现住址"};
        XSSFRow row = sheet.createRow(1);
//        row.setHeightInPoints(31.5F);
        for (int i = 0; i < titles.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(smallRowCellStyle);

        }

        // 设置内容样式
        XSSFCellStyle contentRowCellStyle = workbook.createCellStyle();
        // 1. 克隆已存在的样式--减少代码
        smallRowCellStyle.cloneStyleFrom(titleRowCellStyle);
        // 2. 设置小标题的字体
        XSSFFont contentFront = new XSSFFont();
        contentFront.setFontName("宋体");
        contentFront.setFontHeight((short) 11);
        // 3. 设置不用水平居中
        contentRowCellStyle.setAlignment(HorizontalAlignment.LEFT);
        contentRowCellStyle.setFont(contentFront);


        // 填充内容
        List<User> users = userMapper.selectAll();
        int rowIndex = 2;
        XSSFRow contentRow = null;
        for (User user : users) {
            contentRow = sheet.createRow(rowIndex++);
            XSSFCell cell0 = contentRow.createCell(0);
            cell0.setCellStyle(contentRowCellStyle);
            cell0.setCellValue(user.getId());
            XSSFCell cell1 = contentRow.createCell(1);
            cell1.setCellStyle(contentRowCellStyle);
            cell1.setCellValue(user.getUserName());
            XSSFCell cell2 = contentRow.createCell(2);
            cell2.setCellStyle(contentRowCellStyle);
            cell2.setCellValue(user.getPhone());
            XSSFCell cell3 = contentRow.createCell(3);
            cell3.setCellStyle(contentRowCellStyle);
            cell3.setCellValue(user.getHireDate());
            XSSFCell cell4 = contentRow.createCell(4);
            cell4.setCellStyle(contentRowCellStyle);
            cell4.setCellValue(user.getAddress());
        }
        // 设置文件打开方式
        String filename = "用户表-POI-Excel导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 设置文件类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // 导出
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    /**
     * 通过 Excel 模板进行导出
     *
     * @param response
     */
    public void downLoadXlsxByPoiWithExample(HttpServletResponse response) throws Exception {
        // 1. 获取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); // 获取项目根目录
        File templateFile = new File(rootFile, "/excel_template/userDataExample.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        System.out.println(workbook.getNumberOfSheets());
        // 2. 获取内容单元格格式
        XSSFCellStyle contentRowCellStyle = workbook.getSheetAt(1).getRow(0).getCell(0).getCellStyle();
        // 3. 查找数据
        List<User> users = userMapper.selectAll();
        // 4. 填充单元格
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow contentRow = null;
        int rowIndex = 2;
        for (User user : users) {
            contentRow = sheet.createRow(rowIndex);
            contentRow.setHeightInPoints(16);
            XSSFCell cell0 = contentRow.createCell(0);
            cell0.setCellStyle(contentRowCellStyle);
            cell0.setCellValue(user.getId());
            XSSFCell cell1 = contentRow.createCell(1);
            cell1.setCellStyle(contentRowCellStyle);
            cell1.setCellValue(user.getUserName());
            XSSFCell cell2 = contentRow.createCell(2);
            cell2.setCellStyle(contentRowCellStyle);
            cell2.setCellValue(user.getPhone());
            XSSFCell cell3 = contentRow.createCell(3);
            cell3.setCellStyle(contentRowCellStyle);
            cell3.setCellValue(user.getHireDate());
            XSSFCell cell4 = contentRow.createCell(4);
            cell4.setCellStyle(contentRowCellStyle);
            cell4.setCellValue(user.getAddress());
            rowIndex++;
        }
        workbook.removeSheetAt(1);
        // 5. 导出
        // 设置文件打开方式
        String filename = "用户表-POI-Excel导出By模板.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 设置文件类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // 导出
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    /**
     * 导出某个用户详细信息的Excel
     *
     * @param id
     * @param response
     */
    public void downloadUserInfoByPoiWithTemplate(Long id, HttpServletResponse response) throws Exception {
        // 1. 读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); // 获取项目根目录
        File templateFile = new File(rootFile, "/excel_template/userInfo.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        // 2. 查找用户数据
        User user = userMapper.selectByPrimaryKey(id);
        // 3. 填充用户信息到Excel中
        XSSFSheet sheet = workbook.getSheetAt(0);
        sheet.getRow(1).getCell(1).setCellValue(user.getUserName());
        sheet.getRow(2).getCell(1).setCellValue(user.getPhone());
        sheet.getRow(3).getCell(1).setCellValue(simpleDateFormat.format(user.getBirthday()));
        sheet.getRow(4).getCell(1).setCellValue(user.getSalary());
        sheet.getRow(5).getCell(1).setCellValue(user.getHireDate());
        // 用公式计算司龄 CONCATENATE(DATEDIF(B6,TODAY(),"Y"),"年", DATEDIF(B6,TODAY(),"YM"),"月")
        sheet.getRow(5).getCell(3).setCellFormula("CONCATENATE(DATEDIF(B6,TODAY(),\"Y\"),\"年\", DATEDIF(B6,TODAY(),\"YM\"),\"月\")");
        sheet.getRow(6).getCell(1).setCellValue(user.getProvince());
        sheet.getRow(6).getCell(3).setCellValue(user.getCity());
        sheet.getRow(7).getCell(1).setCellValue(user.getAddress());

        // 填充图片
        // 3.1 先创建一个字节输出流
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        // 3.2 读取图片，放入了一个带有缓存区的图片类中
        BufferedImage bufferedImage = ImageIO.read(new File(rootFile, user.getPhoto()));
        // 计算文件后缀名
        String extName = user.getPhoto().substring(user.getPhoto().lastIndexOf(".") +  1).toUpperCase();
        // 3.3 b把图片写入到字节输出流中
        ImageIO.write(bufferedImage, extName, byteArrayOutputStream);
        // 3.4 Patriarch 控制图片的写入 / ClientAnchor 指定图片的位置
        XSSFDrawing drawingPatriarch = sheet.createDrawingPatriarch();
        // 左上角 偏移x  偏移y  右下角 偏移x 偏移 y  后面就是表格的位置
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 2, 1, 4, 5);
        // 3.5 开始把图片写入到sheet指定的位置
        int format = 0;
        switch (extName){
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


        // 4. 导出
        // 设置文件打开方式
        String filename = "用户" + user.getUserName() + "信息.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 设置文件类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // 导出
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    /**
     * 通过模板引擎动态生成Excel
     * @param id
     * @param response
     */
    public void downloadUserInfoByPoiWithTemplateEngine(Long id, HttpServletResponse response) throws Exception{
        // 1. 读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath()); // 获取项目根目录
        File templateFile = new File(rootFile, "/excel_template/userInfo2.xlsx");
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook(templateFile);
        // 2. 查找用户数据
        User user = userMapper.selectByPrimaryKey(id);
        // 3. 根据引擎填充数据
        workbook = ExcelExportEngine.writeToExcel(user, workbook, rootFile.getPath() + user.getPhoto());
        // 4. 导出
        // 设置文件打开方式
        String filename = "用户" + user.getUserName() + "信息.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        // 设置文件类型
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // 导出
        ServletOutputStream outputStream = response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    /**
     * 百万级数据导出：
     *  1. 肯定得使用高版本的Excel（一个sheet表装的行数多）
     *  2. 使用 sax 方式解析Excel(XML)
     * 限制：
     *  1. 不能使用模板
     *  2. 不能使用太多样式（要不然容易OOM）
     * 注意：
     *  * 这里用的表是 tb_user2，跟之前用的tb_user不一样
     * @param response
     */
    public void downLoadMillion(HttpServletResponse response) throws Exception{
        // 创建使用 sax 方式解析（XML）的workbook
        org.apache.poi.ss.usermodel.Workbook workbook = new SXSSFWorkbook();
        // 分页查询值
        int page = 1;
        // 表示当前数据集合的最后一条id值(也就是代表第几条)
        int num = 0;
        Sheet sheet = null;
        String[] titles = {"编号", "姓名", "手机号", "入职日期", "现住址"};
        int contentIndex = 1;
        Row row = null;
        while(true){
            // 分页查询数据，要不然数据量太多OOM了
            List<User> userList = this.findPage(page++, 10_0000);
            if(CollectionUtils.isEmpty(userList)){
                // 数据已经查完
                break;
            }

            if(num % 100_0000 == 0){
                // 一个sheet存 100_0000条数据，当num % 100_0000后为0，即说明要额外新建sheet表来存储了  0 100_0000 200_0000 300_0000 400_0000
                sheet = workbook.createSheet("第" + (num / 100_0000 + 1) + "个工作表");
                contentIndex = 1;
                Row titleRow = sheet.createRow(0);
                for (int i = 0; i < titles.length; i++) {
                    titleRow.createCell(i).setCellValue(titles[i]);
                }
            }
            for (User user : userList) {
                row = sheet.createRow(contentIndex++);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUserName());
                row.createCell(2).setCellValue(user.getPhone());
                row.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
                row.createCell(4).setCellValue(user.getAddress());
                num++;
            }
            // 4. 导出
            // 设置文件打开方式
            String filename = "百万级用户数据导出.xlsx";
            response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
            // 设置文件类型
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            // 导出
            ServletOutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        }
    }
}
