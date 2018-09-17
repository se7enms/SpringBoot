package com.ms.springboot;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import static org.springframework.boot.SpringApplication.run;

/**
 * @author se7en
 * @date 2018-09-05
 */
@SpringBootApplication
@MapperScan("com.ms.springboot.dao")
public class SpringbootApplication {

	public static void main(String[] args) {
		run(SpringbootApplication.class, args);
	}
}
