/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.utils;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.text.DateFormat;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.Calendar;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * Useful static utilities
 */
@SuppressWarnings("unused")
@NoArgsConstructor(access = AccessLevel.PRIVATE)
@Slf4j
public class Utils {

    public static final String DEFAULT_DATE_FORMAT = "d MMMM yyyy";
    public static final String ALL_NON_NUMERIC_REGEX = "\\D";
    public static final String ALL_NON_FLOAT_REGEX = "[^0-9-. ]";

    private static final String[] DATE_PATTERNS = new String[]{
            "^[a-z]{3} [0-9]{1,2}, [0-9]{4}$",                // 1  'M3 D2, Y4'     12     Jun 21, 1980
            "^[a-z]{3,9} [0-9]{1,2}, [0-9]{4}$",              // 2  'M9 D2, Y4'     18     June 21, 1980
            "^[a-z]{3} [0-9]{1,2} [0-9]{4}$",                 // 3  'M3 D2 Y4'      11     Jun 21 1980
            "^[a-z]{3,9} [0-9]{1,2} [0-9]{4}$",               // 4  'M9 D2 Y4'      17     June 21 1980
            "^[0-9]{1,2}-[a-z]{3,9}-[0-9]{4}$",               // 5  'D2-M3-Y4'      11     21-Jun-1980
            "^[0-9]{1,2}[a-z]{3,9}[0-9]{4}$",                 // 6  'D2M3Y4'         9     21Jun1980
            "^[0-9]{1,2}[a-z]{3,9}[0-9]{2}$",                 // 7  'D2M3Y2'         7     21Jun80
            "^[0-9]{1,2}-[a-z]{3,9}-[0-9]{2}$",               // 8  'D2-M3-Y2'       9     21-Jun-80
            "^[0-9]{1,2} [a-z]{3,9} [0-9]{2}$",               // 9  'D2 M3 Y2'       9     21 Jun 80
            "^[0-9]{1,2} [a-z]{3,9} [0-9]{4}$",               // 10 'D2 M3 Y4'      11     21 Jun 1980
            "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{2}$",               // 11 'M2/D2/Y2'       8     06/21/80
            "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$",               // 12 'M2/D2/Y4'      10     06/21/1980
            "^[a-z]{5,9} [a-z]{3,9} [0-9]{1,2} [0-9]{4}$",    // 13 'W9 M9 D2 Y4'   27     Saturday June 21 1980
            "^[a-z]{5,9} [a-z]{3,9} [0-9]{1,2}, [0-9]{4}$",   // 14 'W9 M9 D2, Y4'  28     Saturday June 21, 1980
            "^[a-z]{5,9} [0-9]{1,2}-[a-z]{3,9}-[0-9]{4}$",    // 15 'W9 D2-M9-Y4'   27     Saturday 21-June-1980
            "^[a-z]{5,9} [0-9]{1,2}-[a-z]{3}-[0-9]{4}$",      // 16 'W9 D2-M3-Y4'   21     Saturday 21-Jun-1980
            "^[a-z]{3} [0-9]{1,2}-[a-z]{3,9}-[0-9]{4}$",      // 17 'W3 D2-M3-Y4'   15     Sat 21-Jun-1980
            "^[a-z]{3} [0-9]{1,2}-[a-z]{3,9}-[0-9]{2}$",      // 18 'W3 D2-M3-Y2'   13     Sat 21-Jun-80
            "^[0-9]{4}-[0-9]{0,3}$",                          // 19 'Y4-J3'          7     1980-173
            "^[0-9]{3,5}$",                                   // 20 'Y2J3'           5     80173
            "^[0-9]{8}$",                                     // 21 'Y4M2D2'         8     19800621
            "^[0-9]{1,2}\\.[0-9]{1,2}\\.[0-9]{2}$",           // 22 'D2.M2.Y2'       8     21.06.80
            "^[0-9]{1,2}\\.[0-9]{1,2}\\.[0-9]{4}$",           // 23 'D2.M2.Y4'      10     21.06.1980
            "^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}$",               // 24 'Y4-M2-D2'      10     1980-06-21
            "^[0-9]{6}$",                                     // 25 'Y2M2D2'         6     800621
            "^[0-9]{1,2}/[a-z]{3,9}/[0-9]{2}$",               // 26 'D2/M3/Y2'       9     21/Jun/80
            "^[0-9]{1,2}/[a-z]{3,9}/[0-9]{4}$"                // 27 'D2/M3/Y4'      11     21/Jun/1980
    };

    private static final String[] TIME_PATTERNS = new String[]{
            "^[0-9]{1,6}$",                                  // 1     'H2M2S2','M2S2','S2'
            "^[0-9]{1,2}.[0-9]{1,2}.[0-9]{1,2}$",            // 2     'H2 M2 S2'
            "^[0-9]{1,2}.[0-9]{1,2}$",                       // 3     'H2 M2'
            "^[0-9]{1,2}\\.[0-9]{1,2} *(am|pm)$",            // 4     'H2.M2 am/pm'
            "^[0-9]{1,2}.[0-9]{1,2}.[0-9]{1,2}\\s+(am|pm)$", // 5     'H2M2S2 am/pm'
            "^[0-9][0-9] [0-9][0-9] [0-9][0-9].[0-9]$"       // 6     'H2 M2 S2.t'
    };

    private static final Map<String, Integer> MONTH_LOOKUPS = new LinkedHashMap<>();

    static {
        MONTH_LOOKUPS.put("january", 0);
        MONTH_LOOKUPS.put("february", 1);
        MONTH_LOOKUPS.put("march", 2);
        MONTH_LOOKUPS.put("april", 3);
        MONTH_LOOKUPS.put("may", 4);
        MONTH_LOOKUPS.put("june", 5);
        MONTH_LOOKUPS.put("july", 6);
        MONTH_LOOKUPS.put("august", 7);
        MONTH_LOOKUPS.put("september", 8);
        MONTH_LOOKUPS.put("october", 9);
        MONTH_LOOKUPS.put("november", 10);
        MONTH_LOOKUPS.put("december", 11);
        MONTH_LOOKUPS.put("jan", 0);
        MONTH_LOOKUPS.put("feb", 1);
        MONTH_LOOKUPS.put("mar", 2);
        MONTH_LOOKUPS.put("apr", 3);
        MONTH_LOOKUPS.put("jun", 5);
        MONTH_LOOKUPS.put("jul", 6);
        MONTH_LOOKUPS.put("aug", 7);
        MONTH_LOOKUPS.put("sep", 8);
        MONTH_LOOKUPS.put("oct", 9);
        MONTH_LOOKUPS.put("nov", 10);
        MONTH_LOOKUPS.put("dec", 11);
    }

    /**
     * Sleeps the current thread for the specified period or until interrupted
     *
     * @param milliseconds Sleep time
     */
    public static void sleep(int milliseconds) {
        try {
            Thread.sleep(milliseconds);
        }
        catch (InterruptedException e) {
            Thread.currentThread().interrupt();        }
    }

    /**
     * Formats a Date object using the pattern DEFAULT_DATE_FORMAT
     *
     * @param date LocalDate object to format
     * @return String
     */
    public static String formatDate(LocalDate date) {
        return formatDate(date, null);
    }

    /**
     * Formats a Date object using the specified pattern
     *
     * @param date   LocalDate object to format
     * @param format Patter - if null, will use default
     * @return String
     */
    public static String formatDate(LocalDate date, String format) {
        if (date == null) {
            return "";
        }
        else if (format == null) {
            format = DEFAULT_DATE_FORMAT;
        }
        return date.format(DateTimeFormatter.ofPattern(format));
    }

    /**
     * Returns a date object from parsing the date string which can be in any one
     * of the standard BASIS formats
     *
     * @param date String form of the date
     * @return LocalDate
     */
    public static LocalDate parseDate(String date) {
        String country = Locale.getDefault().getCountry();

        // The DD/MM/YYYY format is used by most of the world, including Canada whereas
        // on;y a few countries use the MM/DD/YYYY format as default, notably the US

        return parseDate(date, !"US".equalsIgnoreCase(country));
    }

    /**
     * Returns a date object from parsing the date string which can be in any one
     * of the standard BASIS formats. If it can't match any pattern will return null.
     *
     * @param date       String form of the date
     * @param isEuropean If true the parser will use the pattern D2/M2/Y2 and D2/M2/Y4 instead of M2/D2/Y2 and M2/D2/Y2.
     * @return LocalDate
     */
    public static LocalDate parseDate(String date, boolean isEuropean) {
        LocalDate ret = null;
        if (date != null && !date.isEmpty()) {

            // Check the date against the patterns

            Calendar calendar = Calendar.getInstance();
            int century = Calendar.getInstance().get(Calendar.YEAR) - Calendar.getInstance().get(Calendar.YEAR) % 100;
            calendar.clear();
            String[] bits;
            for (int i = 0; i < DATE_PATTERNS.length && !calendar.isSet(Calendar.YEAR); i++) {
                if (date.toLowerCase().matches(DATE_PATTERNS[i])) {
                    switch (i + 1) {
                        case 1:
                        case 2:
                        case 3:
                        case 4:
                            bits = date.split(" |, ");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[1]));
                            calendar.set(Calendar.MONTH, MONTH_LOOKUPS.get(bits[0].substring(0, 3).toLowerCase()));
                            calendar.set(Calendar.YEAR, parseInt(bits[2]));
                            break;
                        case 5:
                        case 8:
                        case 9:
                        case 10:
                        case 26:
                        case 27:
                            bits = date.split("[ -./]");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[0]));
                            calendar.set(Calendar.MONTH, MONTH_LOOKUPS.get(bits[1].substring(0, 3).toLowerCase()));
                            if (bits[2].length() == 2) {
                                calendar.set(Calendar.YEAR, century + parseInt(bits[2]));
                            }
                            else {
                                calendar.set(Calendar.YEAR, parseInt(bits[2]));
                            }
                            break;
                        case 22:
                        case 23:
                            bits = date.split("[ -./]");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[0]));
                            calendar.set(Calendar.MONTH, parseInt(bits[1]) - 1);
                            if (bits[2].length() == 2) {
                                calendar.set(Calendar.YEAR, century + parseInt(bits[2]));
                            }
                            else {
                                calendar.set(Calendar.YEAR, parseInt(bits[2]));
                            }
                            break;
                        case 6:
                        case 7:
                            bits = date.split("[a-zA-Z]{3}");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[0]));
                            if (bits[1].length() == 2) {
                                calendar.set(Calendar.YEAR, century + parseInt(bits[1]));
                            }
                            else {
                                calendar.set(Calendar.YEAR, parseInt(bits[1]));
                            }
                            bits = date.split("\\d[^a-zA-Z]");
                            calendar.set(Calendar.MONTH, MONTH_LOOKUPS.get(bits[1].substring(0, 3).toLowerCase()));
                            break;
                        case 11:
                        case 12:
                            bits = date.split("/");
                            if (isEuropean) {
                                calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[0]));
                                calendar.set(Calendar.MONTH, parseInt(bits[1]) - 1);
                            }
                            else {
                                calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[1]));
                                calendar.set(Calendar.MONTH, parseInt(bits[0]) - 1);
                            }
                            if (bits[2].length() == 2) {
                                calendar.set(Calendar.YEAR, century + parseInt(bits[2]));
                            }
                            else {
                                calendar.set(Calendar.YEAR, parseInt(bits[2]));
                            }
                            break;
                        case 13:
                        case 14:
                            bits = date.split("[ -]|, ");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[2]));
                            calendar.set(Calendar.MONTH, MONTH_LOOKUPS.get(bits[1].substring(0, 3).toLowerCase()));
                            calendar.set(Calendar.YEAR, parseInt(bits[3]));
                            break;
                        case 15:
                        case 16:
                        case 17:
                        case 18:
                            bits = date.split("[ -]|, ");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[1]));
                            calendar.set(Calendar.MONTH, MONTH_LOOKUPS.get(bits[2].substring(0, 3).toLowerCase()));
                            if (bits[3].length() == 2) {
                                calendar.set(Calendar.YEAR, century + parseInt(bits[3]));
                            }
                            else {
                                calendar.set(Calendar.YEAR, parseInt(bits[3]));
                            }
                            break;
                        case 19:
                            bits = date.split("-");
                            calendar.set(Calendar.YEAR, parseInt(bits[0]));
                            calendar.set(Calendar.DAY_OF_YEAR, parseInt(bits[1]));
                            break;
                        case 20:
                            calendar.set(Calendar.YEAR, century + parseInt(date.substring(0, 2)));
                            calendar.set(Calendar.DAY_OF_YEAR, parseInt(date.substring(2)));
                            break;
                        case 21:
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(date.substring(6, 8)));
                            calendar.set(Calendar.MONTH, parseInt(date.substring(4, 6)) - 1);
                            calendar.set(Calendar.YEAR, parseInt(date.substring(0, 4)));
                            break;
                        case 24:
                            bits = date.split("-");
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(bits[2]));
                            calendar.set(Calendar.MONTH, parseInt(bits[1]) - 1);
                            calendar.set(Calendar.YEAR, parseInt(bits[0]));
                            break;
                        case 25:
                            calendar.set(Calendar.DAY_OF_MONTH, parseInt(date.substring(4, 6)));
                            calendar.set(Calendar.MONTH, parseInt(date.substring(2, 4)) - 1);
                            calendar.set(Calendar.YEAR, century + parseInt(date.substring(0, 2)));
                            break;

                        default:
                            break;
                    }
                }
            }

            // If no BASIS pattern was found, then try it against the standard Java
            // otherwise clear the time element from the BASIS date
            if (calendar.isSet(Calendar.YEAR)) {
                return LocalDate.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH) + 1, calendar.get(Calendar.DAY_OF_MONTH));
            }
            else {
                LocalDateTime tmp = parseDateTime(DateFormat.FULL, DateFormat.FULL, date);
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.FULL, DateFormat.LONG, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.FULL, DateFormat.MEDIUM, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.FULL, DateFormat.SHORT, date);
                }

                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.LONG, DateFormat.FULL, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.LONG, DateFormat.LONG, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.LONG, DateFormat.MEDIUM, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.LONG, DateFormat.SHORT, date);
                }

                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.MEDIUM, DateFormat.FULL, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.MEDIUM, DateFormat.LONG, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.MEDIUM, DateFormat.MEDIUM, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.MEDIUM, DateFormat.SHORT, date);
                }

                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.SHORT, DateFormat.FULL, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.SHORT, DateFormat.LONG, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.SHORT, DateFormat.MEDIUM, date);
                }
                if (tmp == null) {
                    tmp = parseDateTime(DateFormat.SHORT, DateFormat.SHORT, date);
                }
                ret = tmp == null ? null : tmp.toLocalDate();
            }
        }
        return ret;
    }

    /**
     * Returns a date object from parsing the time string which can be in any one
     * of the standard BASIS formats
     *
     * @param time String form of the time
     * @return LocalTime
     */
    public static LocalTime parseTime(String time) {

        // Check the date against the patterns
        String timeTest;
        Calendar ret = Calendar.getInstance();
        ret.setTimeInMillis(0);
        if (time != null) {
            time = time.trim().toLowerCase().replace(":", " ");
            for (int iCnt = 0; iCnt < TIME_PATTERNS.length; iCnt++) {
                if (time.toLowerCase().matches(TIME_PATTERNS[iCnt])) {
                    switch (iCnt + 1) {
                        case 1:
                        case 2:
                            time = time.replaceAll(ALL_NON_NUMERIC_REGEX, "");
                            if (time.length() < 6) {
                                time = "000000".substring(time.length()) + time;
                            }
                            ret.set(Calendar.HOUR_OF_DAY, parseInt(time.substring(0, 2)));
                            ret.set(Calendar.MINUTE, parseInt(time.substring(2, 4)));
                            ret.set(Calendar.SECOND, parseInt(time.substring(4, 6)));
                            break;

                        case 3:
                            time = time.replaceAll(ALL_NON_NUMERIC_REGEX, "");
                            if (time.length() < 4) {
                                time = "0000".substring(time.length()) + time;
                            }
                            ret.set(Calendar.HOUR_OF_DAY, parseInt(time.substring(0, 2)));
                            ret.set(Calendar.MINUTE, parseInt(time.substring(2, 4)));
                            break;

                        case 4:
                            timeTest = time.replaceAll(ALL_NON_NUMERIC_REGEX, "").split("(am|pm)")[0].trim();
                            if (timeTest.length() < 4) {
                                timeTest = "0000".substring(timeTest.length()) + timeTest;
                            }
                            ret.set(Calendar.MINUTE, parseInt(timeTest.substring(2, 4)));
                            if (time.endsWith("pm")) {
                                ret.set(Calendar.HOUR_OF_DAY, parseInt(timeTest.substring(0, 2)) + 12);
                            }
                            else {
                                ret.set(Calendar.HOUR_OF_DAY, parseInt(timeTest.substring(0, 2)));
                            }
                            break;
                        case 5:
                            timeTest = time.replaceAll(ALL_NON_NUMERIC_REGEX, "").split("(am|pm)")[0].trim();
                            ret.set(Calendar.MINUTE, parseInt(timeTest.substring(2, 4)));
                            ret.set(Calendar.SECOND, parseInt(timeTest.substring(4, 6)));
                            if (time.endsWith("pm")) {
                                ret.set(Calendar.HOUR_OF_DAY, parseInt(timeTest.substring(0, 2)) + 12);
                            }
                            else {
                                ret.set(Calendar.HOUR_OF_DAY, parseInt(timeTest.substring(0, 2)));
                            }
                            break;
                        case 6:
                            time = time.replaceAll(ALL_NON_NUMERIC_REGEX, "");
                            ret.set(Calendar.HOUR_OF_DAY, parseInt(time.substring(0, 2)));
                            ret.set(Calendar.MINUTE, parseInt(time.substring(2, 4)));
                            ret.set(Calendar.SECOND, parseInt(time.substring(4, 6)));
                            break;

                        default:
                            break;
                    }
                }
            }
            // If no BASIS pattern was found, then try it against the standard Java

            if (!ret.isSet(Calendar.YEAR)) {
                try {
                    ret.setTime(DateFormat.getDateTimeInstance(DateFormat.FULL, DateFormat.FULL).parse(time));
                }
                catch (Exception e) {
                    log.debug(e.getMessage());
                }
            }
        }

        // Clear out the time part of the date
        return LocalTime.of(ret.get(Calendar.HOUR_OF_DAY), ret.get(Calendar.MINUTE), ret.get(Calendar.SECOND));
    }

    /**
     * Attempts to parse a date/time string using the built-in (and
     * rather lame) parse
     *
     * @param dateFormat Style of date format to use
     * @param timeFormat Style of time format to use
     * @param date       Date string to parse
     * @return LocalDateTime object
     */
    private static LocalDateTime parseDateTime(int dateFormat, int timeFormat, String date) {
        LocalDateTime ret = null;
        try {
            return LocalDateTime.ofEpochSecond(DateFormat.getDateTimeInstance(dateFormat, timeFormat).parse(date).getTime() / 1000, 0, ZoneOffset.UTC);
        }
        catch (Exception e) {
            log.debug("Cannot parse date {}", date);
        }
        return ret;
    }

    /**
     * This method parses the date/time according to the specified format
     * and return a Date object
     *
     * @param dateTime Date/Time string
     * @param format   the Date/Time format to be used in the SimpleDateFormat
     * @return Date object
     */
    public static LocalDateTime parseDateTimeInFormat(String dateTime, String format) {
        LocalDateTime ret = null;
        if (dateTime != null && format != null) {
            try {
                DateTimeFormatter dtFormatter = DateTimeFormatter.ofPattern(format);
                return LocalDateTime.parse(dateTime, dtFormatter);
            }
            catch (Exception e) {
                log.debug(e.getMessage());
            }
        }
        return ret;
    }

    /**
     * This method parses the date/time according to the specified format
     * and return a Date object
     *
     * @param dateTime Date/Time string
     * @param format   the Date/Time format to be used in the SimpleDateFormat
     * @return Date object
     */
    public static LocalDate parseDateInFormat(String dateTime, String format) {
        LocalDate ret = null;
        if (dateTime != null && format != null) {
            try {
                DateTimeFormatter dtFormatter = DateTimeFormatter.ofPattern(format);
                return LocalDate.parse(dateTime, dtFormatter);
            }
            catch (Exception e) {
                log.debug(e.getMessage());
            }
        }
        return ret;
    }

    /**
     * Provides a safe way of parsing a string for an Integer without
     * causing the caller too much stress if the string is not actually a number
     *
     * @param value String to convert to a number
     * @return Int Returns 0 if not a number of null
     */
    public static int parseInt(String value) {
        int returnValue = 0;
        if (value != null && !value.isEmpty()) {
            try {
                returnValue = Integer.parseInt(value);
            }
            catch (Exception e) {
                try {
                    returnValue = Integer.parseInt(value.trim().replaceAll(ALL_NON_FLOAT_REGEX, ""));
                }
                catch (Exception e1) {
                    // Nothing to do
                }
            }
        }
        return returnValue;
    }

    /**
     * Joins any objects together using their default toString methods
     * @param objects Array of objects
     * @return String
     */
    public static String join(Object... objects) {
        StringBuilder builder = new StringBuilder();
        if (objects == null) {
            return "";
        }
        for (Object object : objects) {
            if (builder.length() > 0) {
                builder.append(", ");
            }
            builder.append(object);
        }
        return builder.toString();
    }

    /**
     * Returns the number of days between now and the passed date
     *
     * @param since The previous date
     * @return Number of days
     */
    public static long getAge(Instant since) {
        return getAge(since, ChronoUnit.DAYS);
    }

    /**
     * Returns the number of units between now and the passed date
     *
     * @param since The previous date
     * @param unit  Unit type
     * @return Number of units
     */
    public static long getAge(Instant since, ChronoUnit unit) {
        if (since == null) {
            return 0;
        }
        Instant now = Instant.now();
        return unit.between(since.atZone(ZoneOffset.UTC), now.atZone(ZoneOffset.UTC));
    }

    /**
     * Waits for the given pool to exhaust all tasks
     * This is a little special in that in can give feedback
     * as to how the task queue is draining if the executor
     * is of type ThreadPoolExecutor
     *
     * @param executor       Executor to check
     * @param timeoutMinutes Number of minutes to wait for termination
     */
    public static void waitForCompletionWithProgress(ExecutorService executor, int timeoutMinutes) {

        // Shut the executor down so that new jobs cannot be added
        executor.shutdown();

        // Wait for a maximum amount of time
        Instant finish = Instant.now().plus(timeoutMinutes, ChronoUnit.MINUTES);
        try {
            while (!executor.awaitTermination(30, TimeUnit.SECONDS)) {
                if (Instant.now().isAfter(finish)) {
                    log.debug("Waited for {} minutes, but tasks are still running", timeoutMinutes);
                }
                else if (executor instanceof ThreadPoolExecutor) {
                    log.info("Remaining jobs {}", ((ThreadPoolExecutor) executor).getQueue().size());
                }
            }

            // Kill the system now and throw a new exception if we couldn't complete
            if (!executor.isTerminated()) {
                executor.shutdownNow();
                throw new EzGdocs4jException("Waited for %d minutes, but couldn't complete tasks", timeoutMinutes);
            }
        }
        catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new EzGdocs4jException("Task submission got interrupted", e);
        }
    }

    /**
     * Returns true or false if the strings match
     *
     * @param value1 String 1 to match
     * @param values String 2 to match
     *
     * @return boolean
     */
    public static boolean doStringsMatch(String value1, String... values) {
        if (value1 == null && values == null) {
            return true;
        }
        if (value1 == null || values == null) {
            return false;
        }
        for (String value : values) {
            if (value1.equalsIgnoreCase(value)) {
                return true;
            }
        }
        return false;
    }

}
