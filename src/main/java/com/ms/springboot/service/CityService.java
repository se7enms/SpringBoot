package com.ms.springboot.service;

import com.ms.springboot.domain.City;

import java.util.List;

/**
 * 业务逻辑接口类
 *
 * @author se7en
 * @date 2018-09-11
 */

public interface CityService {
    /**
     * 获取所有城市信息
     * @return null
     * @throws Exception error
     */
    List<City> findAllCity() throws Exception;

    /**
     * 根据城市名称，查询
     * @param cityName 城市名
     * @return null
     * @throws Exception error
     */
    City findCityByName(String cityName) throws Exception;

    /**
     * 新增省份城市信息
     * @param city 城市名
     * @throws Exception error
     */
    void saveCityName(City city) throws Exception;

    /**
     * 删除对应省市信息
     * @param ID 序列编号
     * @throws Exception error
     */
    void  deleteCityName(String ID) throws Exception;

}
