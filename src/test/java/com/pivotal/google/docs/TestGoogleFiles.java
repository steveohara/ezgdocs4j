/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.pivotal.google.AbstractTestGoogle;
import com.pivotal.utils.EzGdocs4jException;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 *
 */
@Slf4j
class TestGoogleFiles extends AbstractTestGoogle {

    @Test
    void testCreateAndDelete() {
        try {
            GoogleFile rootFolder = new GoogleFile();
            String ssName = String.format("test-%d", System.currentTimeMillis());
            GoogleSpreadsheet ss = GoogleSpreadsheet.create(ssName);
            assertEquals(ss.getName(), ssName, "New spreadsheet is misnamed");
            List<GoogleFile> files = rootFolder.list((dir, name) -> name.equals(ssName));
            assertFalse(files.isEmpty(), "No new files found");
            ss.delete();
        }
        catch (com.pivotal.google.docs.GoogleException e) {
            throw new EzGdocs4jException(e);
        }
    }

    @Test
    void testCreateInFolderAndDelete() {
        try {
            GoogleFile folder = new GoogleFile();
            String ssName = String.format("test-%d", System.currentTimeMillis());
            GoogleSpreadsheet ss = GoogleSpreadsheet.create(ssName, folder.getFileId());
            assertEquals(ss.getName(), ssName, "New spreadsheet is misnamed");
            List<GoogleFile> files = folder.list((dir, name) -> name.equals(ssName));
            assertFalse(files.isEmpty(), "No new files found");
            ss.delete();
        }
        catch (com.pivotal.google.docs.GoogleException e) {
            throw new EzGdocs4jException(e);
        }
    }

    @Test
    void testCreateFolderFromRoot() {
        try {
            GoogleFile folder = GoogleFile.createFolder("Test Folder for Unit Testing");
            assertTrue(folder.isFolder(), "Folder not correct type");

            GoogleSpreadsheet ss = GoogleSpreadsheet.create("test-spreadsheet");
            ss.move(folder.getFileId());
            List<GoogleFile> files = folder.list();
            assertFalse(files.isEmpty(), "New folder contains files");
            assertTrue(files.get(0).isFile(), "Folder files are not correct");

            for (GoogleFile file : files) {
                file.delete();
            }
            files = folder.list();
            assertTrue(files.isEmpty(), "Cannot delete files");
        }
        catch (com.pivotal.google.docs.GoogleException e) {
            throw new RuntimeException(e);
        }
    }
}
