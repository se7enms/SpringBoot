package com.ms.springboot.service;

import com.ms.springboot.domain.City;

import java.util.List;
import java.util.Map;

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
     * @param cityID ID
     * @return null
     * @throws Exception error
     */
    City findCityByName(Map cityID) throws Exception;

    /**
     * 新增/更新省份城市信息
     * @param city 城市名
     * @throws Exception error
     */
    void saveCityName(Map city) throws Exception;

    /**
     * 新增/更新省份城市信息
     * @param city 城市名
     * @throws Exception error
     */
    void updateCity(Map city) throws Exception;

    /**
     * 删除对应省市信息
     * @param ID 序列编号
     * @throws Exception error
     */
    void  deleteCityName(String ID) throws Exception;

}
