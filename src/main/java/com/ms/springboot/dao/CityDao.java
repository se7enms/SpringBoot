package com.ms.springboot.dao;

import com.ms.springboot.domain.City;
import org.apache.ibatis.annotations.Param;

import java.util.List;

/**
 * dao接口类
 *
 * @author se7en
 * @date 2018-09-11
 */

public interface CityDao {
    /**
     * 根据城市名称，查询城市信息
     *
     * @param cityName 城市名
     * @return null
     */
    City findByName(@Param("cityName") String cityName);

    /**
     * 查找所有城市信息
     * @return List
     */
    List<City> findAllCity();
}
