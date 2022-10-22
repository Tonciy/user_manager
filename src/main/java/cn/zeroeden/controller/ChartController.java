package cn.zeroeden.controller;

import cn.zeroeden.service.ChartService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author Zero
 * @Description 描述此类
 */
@RestController
@RequestMapping("/chart")
public class ChartController {

    @Autowired
    private ChartService chartService;
}
