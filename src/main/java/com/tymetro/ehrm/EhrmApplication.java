package com.tymetro.ehrm;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class EhrmApplication implements CommandLineRunner {
    private static final Logger log = LoggerFactory.getLogger(EhrmApplication.class);

    public static void main(String[] args) {
        SpringApplication.run(EhrmApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
    }
}
