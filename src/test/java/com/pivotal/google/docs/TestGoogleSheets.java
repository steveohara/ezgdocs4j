/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 *
 */
@Slf4j
class TestGoogleSheets {

    @Test
    void testCopySheetFrom() throws GoogleException {
        GoogleSpreadsheet ssFrom = GoogleSpreadsheet.create(String.format("test-%d", System.currentTimeMillis()));
        ssFrom.addSheet("Test");
        assertEquals(2, ssFrom.getSheets().size(), "New sheet creation failed");
        GoogleSpreadsheet ssTo = GoogleSpreadsheet.create(String.format("test-%d", System.currentTimeMillis()));
        ssTo.copyFrom(ssFrom.getSpreadsheetId(), "Test");
        assertEquals(2, ssTo.getSheets().size(), "Copy sheet from failed");
        assertNotNull(ssTo.getSheetByName("Copy of Test"), "Copy sheet from failed with name");
        ssFrom.delete();
        ssTo.delete();
    }

    @Test
    void testCopySheetTo() throws GoogleException {
        GoogleSpreadsheet ssFrom = GoogleSpreadsheet.create(String.format("test-%d", System.currentTimeMillis()));
        ssFrom.addSheet("Test");
        assertEquals(2, ssFrom.getSheets().size(), "New sheet creation failed");
        GoogleSpreadsheet ssTo = GoogleSpreadsheet.create(String.format("test-%d", System.currentTimeMillis()));
        GoogleSheet sheet = ssFrom.getSheetByName("Test");
        sheet.copyTo(ssTo.getSpreadsheetId());
        assertEquals(2, ssTo.getSheets().size(), "Copy sheet to failed");
        assertNotNull(ssTo.getSheetByName("Copy of Test"), "Copy sheet To failed with name");
        ssFrom.delete();
        ssTo.delete();
    }

    @Test
    void testAddToSheet() throws GoogleException {
        GoogleSpreadsheet tmp = GoogleSpreadsheet.create(String.format("test-%d", System.currentTimeMillis()));
        tmp.clear();
        GoogleSheet sheet = tmp.getSheets().get(0);
        sheet.appendValues("A1", "xfvsfvf", "dafdafd");
        sheet.appendValues("A2", "xfvsfvf", "dafdafd");
        sheet.appendValues("A3", "xfvsfvf", "dafdafd");
        sheet.appendValues("A4", "xfvsfvf", "dafdafd");
        sheet.appendValues("A5", "xfvsfvf", "dafdafd");
        sheet.appendValues("A6", "xfvsfvf", "dafdafd");
        tmp.clear();
        tmp.delete();
    }
}
