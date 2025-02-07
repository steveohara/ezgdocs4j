/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google;

import com.google.api.client.http.HttpRequest;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.auth.Credentials;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import com.pivotal.utils.EzGdocs4jException;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

/**
 * Useful common stuff for dealing with the Google API
 */
@Slf4j
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class GoogleServiceFactory {
    public static final String ENV_GOOGLE_CREDENTIALS_FILENAME = "GOOGLE_CREDENTIALS_FILENAME";
    public static final String SYSTEM_GOOGLE_CREDENTIALS_FILENAME = "google.credentials.filename";
    public static final int CONNECT_TIMEOUT = 20000;
    public static final int READ_TIMEOUT = 360000;

    // Efficient multi-threaded HTTP client
    private static final HttpTransport httpTransport = new NetHttpTransport();

    @Getter(AccessLevel.PUBLIC)
    private static Sheets.Spreadsheets sheetsService = null;

    @Getter(AccessLevel.PUBLIC)
    private static Drive.Files filesService = null;

    // Initialise the credentials by looking for the token filename in multiple places
    static {

        // Get the token filename from the known locations, starting with a system variable
        String tokenFilename = System.getProperty(SYSTEM_GOOGLE_CREDENTIALS_FILENAME);
        if (tokenFilename == null || tokenFilename.isEmpty() || !new File(tokenFilename).exists()) {

            // Now try an environment variable
            tokenFilename = System.getenv(ENV_GOOGLE_CREDENTIALS_FILENAME);
            if (tokenFilename == null || tokenFilename.isEmpty() || !new File(tokenFilename).exists()) {

                // Now try a known directory
                String os = System.getProperty("os.name", "generic").toLowerCase(Locale.ENGLISH);
                boolean isWindows = os.contains("win");
                tokenFilename = System.getenv(isWindows ? "USERPROFILE" : "HOME") + (isWindows ? "\\.google\\" : "/.google/") + "credentials.json";
                if (!new File(tokenFilename).exists()) {
                    String error = String.format("Cannot find Google credentials token filename in %s, %s or a file called %s", SYSTEM_GOOGLE_CREDENTIALS_FILENAME, ENV_GOOGLE_CREDENTIALS_FILENAME, tokenFilename);
                    log.error(error);
                    throw new EzGdocs4jException("error");
                }
            }
        }

        // If we have a token filename, then create the required services for all the other
        try {
            List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS, DriveScopes.DRIVE, "https://www.googleapis.com/auth/cloud-billing.readonly");
            GoogleCredentials credentials = GoogleCredentials.fromStream(Files.newInputStream(Paths.get(tokenFilename))).createScoped(scopes);
            HttpRequestInitializer requestInitializer = new LocalHttpCredentialsAdapter(credentials);

            // Create a sheets service to use
            Sheets sheets = new Sheets.Builder(httpTransport, GsonFactory.getDefaultInstance(), requestInitializer).setApplicationName("ezgdocs4j").build();
            sheetsService = sheets.spreadsheets();

            // Create a files service that connects us to the default drive
            Drive drive = new Drive.Builder(httpTransport, GsonFactory.getDefaultInstance(), requestInitializer).setApplicationName("ezgdocs4j").build();
            filesService = drive.files();
        }
        catch (IOException e) {
            log.error("Cannot initialise Google Sheets API service", e);
            throw new EzGdocs4jException(e);
        }
    }

    /**
     * Subclass of the standard Credentials Adapter to allow us to modify the
     * transport timeouts
     */
    private static class LocalHttpCredentialsAdapter extends HttpCredentialsAdapter {
        public LocalHttpCredentialsAdapter(Credentials credentials) {
            super(credentials);
        }

        @Override
        public void initialize(HttpRequest request) throws IOException {
            super.initialize(request);
            request.setConnectTimeout(CONNECT_TIMEOUT);
            request.setReadTimeout(READ_TIMEOUT);
        }
    }

}
