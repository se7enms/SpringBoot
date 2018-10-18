package com.ms.springboot.service.impl;

import com.ms.springboot.dao.CityDao;
import com.ms.springboot.domain.City;
import com.ms.springboot.service.CityService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

/**
 * 业务逻辑实现类
 *
 * @author se7en
 * @date 2018-09-11
 */
@Service
public class CityServiceImpl implements CityService {

    @Autowired
    private CityDao cityDao;

    /**
     * 根据城市名称，查询城市信息
     * @param cityID ID
     * @return 城市名
     */
    @Override
    public City findCityByName(Map cityID) {
        return cityDao.findByName(cityID);
    }

    /**
     * 新增省市信息
     * @param city 城市名
     */
    @Override
    public void saveCityName(Map city) {
        cityDao.saveCity(city);
    }

    /**
     * 更新省市信息
     * @param city 城市名
     */
    @Override
    public void updateCity(Map city) {
        cityDao.updateCity(city);
    }

    /**
     * 删除城市信息
     * @param ID 数据ID
     */
    @Override
    public void deleteCityName(String ID) {
        cityDao.deleteCityName(ID);
    }

    /**
     * 获取所有城市信息
     * @return List
     */
    @Override
    public List<City> findAllCity() {
        return cityDao.findAllCity();
    }
}
