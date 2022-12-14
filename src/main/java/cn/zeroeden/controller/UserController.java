package cn.zeroeden.controller;

import cn.zeroeden.pojo.User;
import cn.zeroeden.service.UserService;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.category.DefaultCategoryDataset;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.File;
import java.util.List;

@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private UserService userService;

    @GetMapping("/findPage")
    public List<User> findPage(
            @RequestParam(value = "page", defaultValue = "1") Integer page,
            @RequestParam(value = "rows", defaultValue = "10") Integer pageSize) {
        return userService.findPage(page, pageSize);
    }

    @GetMapping(value = "downLoadXlsByJxl", name = "通过jxl框架导出Excel")
    public void downLoadXlsByJxl(HttpServletResponse response) throws Exception {
        userService.downLoadXlsByJxl(response);
    }

    @PostMapping(value = "uploadExcel", name = "上传Excel数据")
    public void uploadExcel(MultipartFile file) throws Exception {
        userService.uploadExcel(file);
    }

    @GetMapping(value = "downLoadXlsxByPoi", name = "通过POI下载Excel")
    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
//        userService.downLoadXlsxByPoi(response);
        // 带样式的Excel下载
//        userService.downLoadXlsxByPoiWithStyle(response);

        // 使用模板生成导出Excel
        userService.downLoadXlsxByPoiWithExample(response);
    }

    @GetMapping(value = "download", name = "通过POI下载用户详细信息Excel")
    public void downloadUserInfoByPoiWithTemplate(Long id, HttpServletResponse response) throws Exception {
//        通过模板来导出数据
//        userService.downloadUserInfoByPoiWithTemplate(id, response);
        // 通过模板引擎来动态性导出数据
        userService.downloadUserInfoByPoiWithTemplateEngine(id, response);
    }


    /**
     * 根据模板引擎来动态性导出Excel
     *
     * @param id
     * @param response
     * @throws Exception
     */
    public void downloadUserInfoByPoiWithTemplateEngine(Long id, HttpServletResponse response) throws Exception {
        userService.downloadUserInfoByPoiWithTemplateEngine(id, response);
    }


    @GetMapping(value = "downLoadMillion", name = "利用POI（sax）导出百万级数据")
    public void downLoadMillion(HttpServletResponse response) throws Exception {
        userService.downLoadMillion(response);
    }

    @GetMapping(value = "downLoadCSV", name = "通过CSV文件装载百万级数据")
    public void downLoadCSV(HttpServletResponse response) throws Exception {
        userService.downLoadCSV(response);
    }

    @GetMapping(value = "/{id}", name = "通过Id查找用户信息")
    public User findById(@PathVariable("id") Long id) throws Exception {
        return userService.findById(id);
    }

    @GetMapping(value = "/downloadContract", name = "通过Id导出用户合同word文档")
    public void downloadContract(Long id, HttpServletResponse response) throws Exception {
        userService.downloadContract(id, response);
    }

    @GetMapping(value = "/downLoadWithEasyPOI", name = "通过EasyPOI框架导出Excel")
    public void downLoadWithEasyPOI(HttpServletResponse response) throws Exception {
        userService.downLoadWithEasyPOI(response);
    }

    @GetMapping(value = "/jfreeChart", name = "通过JFreeChart框架导出图片")
    public void jfreeChart(HttpServletResponse response) throws Exception {
        // 1. 装备数据集
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        // y值，折线名称，x值
        dataset.addValue(200, "公司", "2021年");
        dataset.addValue(250, "公司", "2022年");
        dataset.addValue(100, "公司", "2023年");
        dataset.addValue(400, "公司", "2024年");
        dataset.addValue(200, "企业", "2021年");
        dataset.addValue(250, "企业", "2022年");
        dataset.addValue(100, "企业", "2023年");
        dataset.addValue(400, "企业", "2024年");
        // 构造图表的主题样式
        StandardChartTheme chartTheme = new StandardChartTheme("CN");
        // 设置大标题的字体
        chartTheme.setExtraLargeFont(new Font("宋体",Font.BOLD, 20));
        // 设置图例的字体
        chartTheme.setRegularFont(new Font("宋体", Font.BOLD, 15));
        // 设置内容及x/y轴的字体
        chartTheme.setLargeFont(new Font("宋体", Font.BOLD, 15));
        ChartFactory.setChartTheme(chartTheme);
        // 2. 构造折现图  大标题  x轴说明   y轴说明    数据集
        JFreeChart chart = ChartFactory.createBarChart("入职人数","年份","人数", dataset);
        // 3. 生成图--放到输出流中
        ChartUtils.writeChartAsJPEG(response.getOutputStream(), chart, 600, 400);
    }



}
