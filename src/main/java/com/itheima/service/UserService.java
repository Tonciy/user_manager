package com.itheima.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.ibatis.annotations.Case;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class UserService {

    @Autowired
    private UserMapper userMapper;

    // 处理日期转化
    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page,pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws  Exception{
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
                switch (i){
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
}
