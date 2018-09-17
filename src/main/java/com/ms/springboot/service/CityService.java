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
     * 根据城市名称，查询城市信息
     * @param cityName 城市名
     * @return null
     * @throws Exception error
     */
    City findCityByName(String cityName) throws Exception;

    /**
     * 获取所有城市信息
     * @return null
     * @throws Exception error
     */
    List<City> findAllCity() throws Exception;
}
