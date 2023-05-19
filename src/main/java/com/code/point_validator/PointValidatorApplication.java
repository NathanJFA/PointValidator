package com.code.point_validator;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PointValidatorApplication implements CommandLineRunner {
	@Autowired
	ComparatorSheets comparatorSheets;
	public static void main(String[] args) {
		SpringApplication.run(PointValidatorApplication.class, args);
	}
	@Override
	public void run(String... args) throws Exception {
		comparatorSheets.sysMain(args[0], args[1]);
	}
}
