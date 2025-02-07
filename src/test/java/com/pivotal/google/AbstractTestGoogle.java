package com.pivotal.google;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.BeforeAll;

import java.io.IOException;
import java.util.logging.LogManager;

@Slf4j
public class AbstractTestGoogle {

    public static final String SIMPLE_LOGGING_PROPERTIES_FILENAME = "/logging.properties";

    @BeforeAll
    public static void setUpLogger() {

        // Setup the Java simple logger
        LogManager logManager = LogManager.getLogManager();
        try {
            logManager.readConfiguration(GoogleServiceFactory.class.getResourceAsStream(SIMPLE_LOGGING_PROPERTIES_FILENAME));
        }
        catch (IOException e) {
            log.warn("Cannot load simple log configuration file [{}] from classpath  - {}", SIMPLE_LOGGING_PROPERTIES_FILENAME, e.getMessage());
        }
    }

}
