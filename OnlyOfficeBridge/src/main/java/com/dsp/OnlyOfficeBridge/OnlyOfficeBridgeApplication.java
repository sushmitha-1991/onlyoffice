package com.dsp.OnlyOfficeBridge;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableAsync;

@SpringBootApplication
@EnableAsync
public class OnlyOfficeBridgeApplication {

	public static void main(String[] args) {
		SpringApplication.run(OnlyOfficeBridgeApplication.class, args);
	}

}
