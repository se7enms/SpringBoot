package com.ms.springboot.controller;

import com.ms.springboot.domain.City;
import com.ms.springboot.service.CityService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
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
     * @return html
     * @throws Exception error
     */
    @RequestMapping(value = "/city", method = RequestMethod.GET)
    public String getCityList(ModelMap map) throws Exception {
        map.addAttribute("cityList",cityService.findAllCity());
        return BOOK_LIST_PATH_NAME;
    }

    /**
     * 新增一条城市信息
     * @param city 城市信息集合
     * @return html
     * @throws Exception error
     */
    @RequestMapping(value = "/createCity", method = RequestMethod.POST)
    public String postCity(@ModelAttribute City city) throws Exception {
        cityService.saveCityName(city);
        return BOOK_LIST_PATH_NAME;
    }

    /**
     * 删除城市信息
     * @param id 信息列表ID
     * @return html
     * @throws Exception error
     */
    @RequestMapping(value = "/delCity/{id}", method = RequestMethod.GET)
    public String deleteCity(@PathVariable String id) throws Exception {
        System.out.print(id);
        cityService.deleteCityName(id);
        //重定向，刷新页面
        return "redirect:/index/city";
    }
}
