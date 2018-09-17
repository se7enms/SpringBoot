package com.ms.springboot.controller;

import com.ms.springboot.service.CityService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * My First Demo
 * @author se7en
 * @date 2018-09-05
 */
@Controller
@RequestMapping(value = "/index")
public class CityController {
    @Autowired
    private CityService cityService;

    private static final String BOOK_LIST_PATH_NAME = "cityList";

    /**
     * 查询所有城市
     * @return map
     * @throws Exception error
     */
    @RequestMapping(value = "/city", method = RequestMethod.GET)
    public String getCityList(ModelMap map) throws Exception {
        map.addAttribute("cityList",cityService.findAllCity());
        return BOOK_LIST_PATH_NAME;
    }
}
