/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.google.api.services.sheets.v4.model.GridRange;
import com.pivotal.utils.EzGdocs4jException;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 *
 */
@Slf4j
class TestGoogleServiceFactory {

    @Test
    void testR1C1Conversion() {
        try {
            assertEquals("9,0,10,1", getRangeString(GoogleDocsUtils.getGridRange(1, "A10")), "Incorrect coordinates");
            assertEquals("null,0,null,2", getRangeString(GoogleDocsUtils.getGridRange(1, "A:B")), "Incorrect coordinates");
            assertEquals("0,0,1,1", getRangeString(GoogleDocsUtils.getGridRange(1, "A1")), "Incorrect coordinates");
            assertEquals("null,51,null,52", getRangeString(GoogleDocsUtils.getGridRange(1, "AZ")), "Incorrect coordinates");
            assertEquals("null,26,null,54", getRangeString(GoogleDocsUtils.getGridRange(1, "AA:BB")), "Incorrect coordinates");
            assertEquals("0,2,null,4", getRangeString(GoogleDocsUtils.getGridRange(1, "C1:D")), "Incorrect coordinates");
            assertEquals("0,2,5,4", getRangeString(GoogleDocsUtils.getGridRange(1, "C1:D5")), "Incorrect coordinates");
            assertThrows(com.pivotal.google.docs.GoogleException.class, () -> GoogleDocsUtils.getGridRange(1, "C:^"));
            assertThrows(com.pivotal.google.docs.GoogleException.class, () -> GoogleDocsUtils.getGridRange(1, "1C"));
        }
        catch (com.pivotal.google.docs.GoogleException e) {
            throw new EzGdocs4jException(e);
        }

    }

    private String getRangeString(GridRange range) {
        return range.getStartRowIndex() + "," +
                range.getStartColumnIndex() + "," +
                range.getEndRowIndex() + "," +
                range.getEndColumnIndex();
    }

}
