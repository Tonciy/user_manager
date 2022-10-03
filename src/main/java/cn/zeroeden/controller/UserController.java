package cn.zeroeden.controller;

import cn.zeroeden.pojo.User;
import cn.zeroeden.service.UserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
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

    @GetMapping(value = "downLoadXlsxByPoi",name = "通过POI下载Excel")
    public  void downLoadXlsxByPoi(HttpServletResponse response) throws Exception{
//        userService.downLoadXlsxByPoi(response);
        // 带样式的Excel下载
//        userService.downLoadXlsxByPoiWithStyle(response);

        // 使用模板生成导出Excel
        userService.downLoadXlsxByPoiWithExample(response);
    }

    @GetMapping(value = "download",name = "通过POI下载用户详细信息Excel")
    public  void downloadUserInfoByPoiWithTemplate(Long id, HttpServletResponse response) throws Exception{
//        通过模板来导出数据
//        userService.downloadUserInfoByPoiWithTemplate(id, response);
        // 通过模板引擎来动态性导出数据
        userService.downloadUserInfoByPoiWithTemplateEngine(id, response);
    }


    /**
     * 根据模板引擎来动态性导出Excel
     * @param id
     * @param response
     * @throws Exception
     */
    public  void downloadUserInfoByPoiWithTemplateEngine(Long id, HttpServletResponse response) throws Exception{
        userService.downloadUserInfoByPoiWithTemplateEngine(id, response);
    }


    @GetMapping(value = "downLoadMillion",name = "利用POI（sax）导出百万级数据")
    public  void downLoadMillion(HttpServletResponse response) throws Exception{
        userService.downLoadMillion( response);
    }



}
