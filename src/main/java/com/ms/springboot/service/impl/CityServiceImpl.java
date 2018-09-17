package com.ms.springboot.service.impl;

import com.ms.springboot.dao.CityDao;
import com.ms.springboot.domain.City;
import com.ms.springboot.service.CityService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

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

    @Override
    public City findCityByName(String cityName) {
        return cityDao.findByName(cityName);
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
