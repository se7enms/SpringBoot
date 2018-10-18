package com.ms.springboot.controller;

import com.ms.springboot.service.CityService;
import net.sf.json.JSONArray;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.Map;

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

    private static final String CITY_LIST_PATH_NAME = "cityList";
    private static final String CITY_FORM_PATH_NAME = "cityForm";

    /**
     * 查询所有城市
     * @return html
     * @throws Exception error
     */
    @RequestMapping(value = "/city")
    public String getCityList(ModelMap map) throws Exception {
        map.addAttribute("cityList",cityService.findAllCity());
        return CITY_LIST_PATH_NAME;
    }

    /**
     * 打开新增/编辑页面
     * @param model 城市信息
     * @return html
     */
    @RequestMapping(value = "/toAddCity", method = RequestMethod.GET)
    public String postCity(Model model,HttpServletRequest request) throws Exception {
        Map<String, String> map = new HashMap<>(16);
        String addValue = "add";
        String updateValue = "update";
        String state = request.getParameter("state");
        if(addValue.equals(state)) {
            //新增城市
            model.addAttribute("action", "addCity");
        } else if (updateValue.equals(state)) {
            //获取编辑城市的信息
            String ID = request.getParameter("ID");
            map.put("ID",ID);
            model.addAttribute("cityInfo", cityService.findCityByName(map));
            model.addAttribute("ID", ID);
            model.addAttribute("action", "updateCity");
        }
        return CITY_FORM_PATH_NAME;
    }

    /**
     * 保存城市信息
     * @throws Exception error
     */
    @RequestMapping(value = "/addCity", method = RequestMethod.POST)
    public void saveCity(HttpServletRequest request, HttpServletResponse response) throws Exception {
        Map<String, String> map = new HashMap<>(16);
        String addValue = "addCity";
        String updateValue = "updateCity";

        String provinceName = request.getParameter("provinceName");
        String cityName = request.getParameter("cityName");
        String description = request.getParameter("description");
        String state = request.getParameter("state");
        map.put("provinceName",provinceName);
        map.put("cityName",cityName);
        map.put("description",description);

        if (addValue.equals(state)) {
            cityService.saveCityName(map);
        } else if (updateValue.equals(state)) {
            map.put("ID", request.getParameter("ID"));
            cityService.updateCity(map);
        }

        JSONArray jsonArray = JSONArray.fromObject(map);
        try {
            response.setContentType("text/html;charset=UTF-8");
            PrintWriter out = response.getWriter();
            out.print(jsonArray.toString());
            out.flush();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 删除城市信息
     * @param id 信息列表ID
     * @return html
     * @throws Exception error
     */
    @RequestMapping(value = "/delCity/{id}", method = RequestMethod.GET)
    public String deleteCity(@PathVariable String id) throws Exception {
        cityService.deleteCityName(id);
        //重定向，刷新页面
        return "redirect:/index/city";
    }
}
