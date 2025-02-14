/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.services.sheets.v4.model.*;
import com.pivotal.google.GoogleServiceFactory;
import com.pivotal.utils.Utils;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Useful common stuff for dealing with the Google API
 */
@Slf4j
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class GoogleDocsUtils {
    // Precomputed difference between the Unix epoch and the Google Sheets epoch
    private static final long SHEETS_EPOCH_DIFFERENCE = 2209161600000L;
    public static final int MAX_RATE_LIMIT_RETRIES = 10;
    public static final int GAPI_RETRY_ERROR_CODE = 429;
    public static final double MAXIMUM_RETRY_SLEEP = 32000.0;

    /**
     * Useful set of criteria types when building FilterCriteria
     * @noinspection unused
     */
    public enum BooleanCriteriaType {
        CELL_EMPTY, // The criteria is met when a cell is empty.
        CELL_NOT_EMPTY, // The criteria is met when a cell is not empty.
        DATE_AFTER, // The criteria is met when a date is after the given value.
        DATE_BEFORE, // The criteria is met when a date is before the given value.
        DATE_EQUAL_TO, // The criteria is met when a date is equal to the given value.
        DATE_NOT_EQUAL_TO, // The criteria is met when a date is not equal to the given value.
        DATE_AFTER_RELATIVE, // The criteria is met when a date is after the relative date value.
        DATE_BEFORE_RELATIVE, // The criteria is met when a date is before the relative date value.
        DATE_EQUAL_TO_RELATIVE, // The criteria is met when a date is equal to the relative date value.
        NUMBER_BETWEEN, // The criteria is met when a number that is between the given values.
        NUMBER_EQUAL_TO, // The criteria is met when a number that is equal to the given value.
        NUMBER_GREATER_THAN, // The criteria is met when a number that is greater than the given value.
        NUMBER_GREATER_THAN_OR_EQUAL_TO, // The criteria is met when a number that is greater than or equal to the given value.
        NUMBER_LESS_THAN, // The criteria is met when a number that is less than the given value.
        NUMBER_LESS_THAN_OR_EQUAL_TO, // The criteria is met when a number that is less than or equal to the given value.
        NUMBER_NOT_BETWEEN, // The criteria is met when a number that is not between the given values.
        NUMBER_NOT_EQUAL_TO, // The criteria is met when a number that is not equal to the given value.
        TEXT_CONTAINS, // The criteria is met when the input contains the given value.
        TEXT_DOES_NOT_CONTAIN, // The criteria is met when the input does not contain the given value.
        TEXT_EQ, // The criteria is met when the input is equal to the given value.
        TEXT_NOT_EQ, // The criteria is met when the input is not equal to the given value.
        TEXT_STARTS_WITH, // The criteria is met when the input begins with the given value.
        TEXT_ENDS_WITH, // The criteria is met when the input ends with the given value.
        CUSTOM_FORMULA // The criteria is met when the input makes the given formula evaluate to true.
    }

    /**
     * Takes the requests and executes them as a batch
     *
     * @param spreadsheetId Spreadsheet to work on
     * @param request       Requests (can be multiple)
     * @return BatchUpdateSpreadsheetResponse Response to get returned values
     * @throws GoogleException If the batch fails
     */
    public static BatchUpdateSpreadsheetResponse executeBatchRequest(String spreadsheetId, Request... request) throws GoogleException {
        BatchUpdateSpreadsheetRequest batchRequest = new BatchUpdateSpreadsheetRequest();
        batchRequest.setRequests(Arrays.asList(request));

        // We are going to retry up to 5 times with an exponential time based back-off
        // if we get the dreaded rate limits
        int retries = 0;
        while (true) {
            try {
                return GoogleServiceFactory.getSheetsService().batchUpdate(spreadsheetId, batchRequest).execute();
            }
            catch (IOException e) {
                retries = sleepWithBackOff(e, retries);
            }
        }
    }

    /**
     * Returns a spreadsheet by retrieving it from the server
     *
     * @param spreadsheetId ID of the spreadsheet
     * @return Spreadsheet object
     * @throws GoogleException If the spreadsheet cannot be found/opened
     */
    protected static Spreadsheet getSpreadsheet(String spreadsheetId) throws GoogleException {
        int retries = 0;
        while (true) {
            try {
                return GoogleServiceFactory.getSheetsService().get(spreadsheetId).execute();
            }
            catch (IOException e) {
                retries = sleepWithBackOff(e, retries);
            }
        }
    }

    /**
     * Used to do a managed sleep to exponentially back off from sending
     * more requests if we are exceeding the rate limit
     *
     * @param e       Exception thrown by the
     * @param retries Number of retries we have made
     * @return Updated retries count
     * @throws GoogleException If the exception is mot rate limited or we have exceeded our retries
     */
    public static int sleepWithBackOff(IOException e, int retries) throws GoogleException {
        retries++;
        if (retries > MAX_RATE_LIMIT_RETRIES) {
            throw new GoogleException("Rate limit retry attempts exceeded", e);
        }

        // Check if this is actually a proper response
        if (e instanceof GoogleJsonResponseException) {
            if (((GoogleJsonResponseException) e).getDetails().getCode() == GAPI_RETRY_ERROR_CODE) {
                log.warn("Retrying API request after a rate limit error [{} of {}]", retries, MAX_RATE_LIMIT_RETRIES);
                double sleep = Math.min((Math.pow(2, retries) + Math.random()) * 1000, MAXIMUM_RETRY_SLEEP);
                Utils.sleep((int) sleep);
            }
            else {
                throw new GoogleException("Cannot run Google API request [%s]\n%s", e, e.getMessage(), e.getStackTrace());
            }
        }
        else {
            throw new GoogleException("Request failed through a network error - {}", e, e.getMessage());
        }
        return retries;
    }

    /**
     * Retrieves the sheet by name by first getting the spreadsheet from the server
     *
     * @param spreadsheetId ID of the spreadsheet
     * @param sheetId       Name of the sheet
     * @return A Sheet object if the sheet exists, null otherwise
     * @throws GoogleException If the spreadsheet cannot be found/opened
     */
    protected static Sheet getSheetById(String spreadsheetId, int sheetId) throws GoogleException {
        List<Sheet> sheets = getSpreadsheet(spreadsheetId).getSheets();
        if (sheets != null) {
            for (Sheet sheet : sheets) {
                if (sheet.getProperties().getSheetId() == sheetId) {
                    return sheet;
                }
            }
        }
        return null;
    }

    /**
     * Retrieves thw sheet by name by first getting the spreadsheet from the server
     *
     * @param spreadsheetId ID of the spreadsheet
     * @param name          Name of the sheet
     * @return A Sheet object if the sheet exists, null otherwise
     * @throws GoogleException If the spreadsheet cannot be found/opened
     */
    protected static Sheet getSheetByName(String spreadsheetId, String name) throws GoogleException {
        List<Sheet> sheets = getSpreadsheet(spreadsheetId).getSheets();
        if (sheets != null) {
            for (Sheet sheet : sheets) {
                if (sheet.getProperties().getTitle().equalsIgnoreCase(name)) {
                    return sheet;
                }
            }
        }
        return null;
    }

    /**
     * Convenience method for converting a standard Java Color into a Google Color
     *
     * @param color AWT Color
     * @return Google Color
     */
    protected static Color convertColorToGoogle(java.awt.Color color) {
        Color googleColor = new Color();
        if (color != null) {
            float[] components = color.getRGBComponents(null);
            googleColor.setRed(components[0]);
            googleColor.setGreen(components[1]);
            googleColor.setBlue(components[2]);
            googleColor.setAlpha(components[3]);
        }
        return googleColor;
    }

    /**
     * Gets a repeating cell request to update some contiguous cells
     *
     * @param sheetId     Sheet ID to assign
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format (must be greater than startRow)
     * @param endColumn   End column of range to format (must be greater than endColumn)
     * @return RepeatCellRequest used to update multiple contiguous cells
     */
    protected static RepeatCellRequest getRepeatCellRequest(int sheetId, Integer startRow, Integer startColumn, Integer endRow, Integer endColumn) {
        RepeatCellRequest req = new RepeatCellRequest();
        GridRange gridRange = getGridRange(sheetId, startRow, startColumn, endRow, endColumn);
        req.setRange(gridRange);
        return req;
    }

    /**
     * Converts a LocalDate date to an epoch date value suitable for Google sheets
     *
     * @param date LocalDate date
     * @return String epoch date
     */
    public static double getGoogleDateValue(LocalDate date) {
        if (date == null) {
            return 0;
        }
        long millisSinceUnixEpoch = date.toEpochDay() * TimeUnit.DAYS.toMillis(1);
        long millisSinceSheetsEpoch = millisSinceUnixEpoch + SHEETS_EPOCH_DIFFERENCE;
        return millisSinceSheetsEpoch / (double) TimeUnit.DAYS.toMillis(1);
    }

    /**
     * Gets a grid using the specified coordinates
     *
     * @param sheetId     Sheet ID to assign
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format (must be greater than startRow)
     * @param endColumn   End column of range to format (must be greater than endColumn)
     * @return GridRange used to update multiple contiguous cells
     */
    protected static GridRange getGridRange(int sheetId, Integer startRow, Integer startColumn, Integer endRow, Integer endColumn) {

        // Set the range to affect
        GridRange gridRange = new GridRange();
        if (startRow != null) {
            gridRange.setStartRowIndex(startRow);
        }
        if (startColumn != null) {
            gridRange.setStartColumnIndex(startColumn);
        }
        if (endRow != null) {
            gridRange.setEndRowIndex(endRow);
        }
        if (endColumn != null) {
            gridRange.setEndColumnIndex(endColumn);
        }
        gridRange.setSheetId(sheetId);
        return gridRange;
    }

    /**
     * Gets a grid using the specified coordinates
     *
     * @param sheetId    Sheet ID to assign
     * @param a1Notation Range specified as A1 type notation
     * @return GridRange used to update multiple contiguous cells
     * @throws GoogleException if the range specification is invalid
     */
    public static GridRange getGridRange(int sheetId, String a1Notation) throws GoogleException {

        // Check the notation looks OK
        if (a1Notation == null || a1Notation.isEmpty()) {
            throw new GoogleException("A1Notation is required");
        }
        a1Notation = a1Notation.replaceAll("(?i)[^A-Z:0-9]", "");
        if (!a1Notation.matches("(?i)[A-Z]+(\\d+)?(:[A-Z]+(\\d+)?)?")) {
            throw new GoogleException("A1Notation is required");
        }

        Integer startRow = null;
        Integer startColumn = null;
        Integer endRow = null;
        Integer endColumn = null;

        // Get the first part
        String part = a1Notation.contains(":") ? a1Notation.split(":", 2)[0] : a1Notation;
        String col = part.replaceAll("(?i)[^A-Z]", "");
        if (col.isEmpty()) {
            throw new GoogleException("A1Notation %s is invalid", a1Notation);
        }
        startColumn = parseANotationString(col);
        String row = part.replaceAll("(?i)\\D", "");
        if (!row.isEmpty()) {
            startRow = Integer.parseInt(row) - 1;
        }

        // Get the second part
        if (a1Notation.contains(":")) {
            part = a1Notation.split(":", 2)[1];
            col = part.replaceAll("(?i)[^A-Z]", "");
            if (col.isEmpty()) {
                throw new GoogleException("A1Notation %s is invalid", a1Notation);
            }
            endColumn = parseANotationString(col) + 1;
            row = part.replaceAll("(?i)\\D", "");
            if (!row.isEmpty()) {
                endRow = Integer.parseInt(row);
            }
        }

        // No second part, so adjust the range
        else {
            endColumn = startColumn + 1;
            if (startRow != null) {
                endRow = startRow + 1;
            }
        }

        // Set the range to affect
        log.debug("Converted [{}] into {},{},{},{}", a1Notation, startRow, startColumn, endRow, endColumn);
        return getGridRange(sheetId, startRow, startColumn, endRow, endColumn);
    }

    /**
     * Turns the column specification into a numeric value e.g. AZ == 51
     *
     * @param val String in normal A1 type notation e.g. 'A', 'AAZ' etc.
     * @return Numeric column number
     */
    private static int parseANotationString(String val) {
        int ret = 0;
        if (val != null) {
            val = val.toLowerCase();
            for (int i = 0; i < val.length(); i++) {
                if (i > 0) {
                    ret = (ret + 1) * 26;
                }
                ret += val.charAt(i) - 'a';
            }
        }
        return ret;
    }
}
