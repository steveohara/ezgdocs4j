/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.pivotal.google.GoogleServiceFactory;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.*;

/**
 * Provides a means to work with the rather horrible Google Sheets API
 * @noinspection UnusedReturnValue
 */
@SuppressWarnings("unused")
@Slf4j
@Getter
@Setter
public class GoogleSpreadsheet {

    private String spreadsheetId;

    /**
     * Opens a spreadsheet using the specified ID
     *
     * @param spreadsheetId Unique ID
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSpreadsheet(String spreadsheetId) throws GoogleException {
        try {
            GoogleServiceFactory.getSheetsService().get(spreadsheetId).execute();
            this.spreadsheetId = spreadsheetId;
        }
        catch (IOException e) {
            throw new GoogleException("Cannot open spreadsheet %s", e, spreadsheetId);
        }
    }

    /**
     * Creates a new spreadsheet in the root folder
     * If the spreadsheet already exists, then it simply returns it
     *
     * @param name Name to give the spreadsheet
     * @return GoogleSpreadsheet to continue working with it
     * @throws GoogleException If there is some sort of error
     */
    public static GoogleSpreadsheet create(String name) throws GoogleException {
        return create(name, null);
    }

    /**
     * Creates a new spreadsheet in the given folder
     * If the spreadsheet already exists, then it simply returns it
     *
     * @param name     Name to give the spreadsheet
     * @param folderId Drive folder to create it in
     * @return GoogleSpreadsheet to continue working with it
     * @throws GoogleException If there is some sort of error
     */
    public static GoogleSpreadsheet create(String name, String folderId) throws GoogleException {

        // Create the spreadsheet in the ROOT folder
        Spreadsheet requestBody = new Spreadsheet();
        SpreadsheetProperties properties = new SpreadsheetProperties();
        properties.setTitle(name);
        requestBody.setProperties(properties);
        try {
            Sheets.Spreadsheets.Create req = GoogleServiceFactory.getSheetsService().create(requestBody);
            Spreadsheet spreadsheet = req.setFields("spreadsheetId").execute();
            GoogleSpreadsheet ret = new GoogleSpreadsheet(spreadsheet.getSpreadsheetId());

            // Now attempt to move the sheet to the folder
            if (folderId != null && !folderId.isEmpty()) {
                try {
                    ret.move(folderId);
                }
                catch (GoogleException e) {
                    log.warn("Cannot open destination folder {}", folderId);
                }
            }
            return ret;
        }
        catch (IOException e) {
            throw new GoogleException("Cannot create spreadsheet %s", e, name);
        }
    }

    /**
     * Copies the sheet if it exists from the specified spreadsheet
     * If the sheet doesn't exist, returns null otherwise it returns the new sheet
     *
     * @param fromSpreadsheetId Spreadsheet to copy from
     * @param sheetName         Name of the sheet to copy
     * @return The duplicated sheet or null if it doesn't exist
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet copyFrom(String fromSpreadsheetId, String sheetName) throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetByName(fromSpreadsheetId, sheetName);
        if (sheet != null) {
            return copyFrom(fromSpreadsheetId, sheet.getProperties().getSheetId());
        }
        return null;
    }

    /**
     * Copies the sheet if it exists from the specified spreadsheet
     * If the sheet doesn't exist, returns null otherwise it returns the new sheet
     *
     * @param fromSpreadsheetId Spreadsheet to copy from
     * @param sheetId           ID of the sheet to copy
     * @return The duplicated sheet or null if it doesn't exist
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet copyFrom(String fromSpreadsheetId, int sheetId) throws GoogleException {

        // Check if the sheet already exists
        Sheet sheet = GoogleDocsUtils.getSheetById(fromSpreadsheetId, sheetId);
        if (sheet != null) {
            CopySheetToAnotherSpreadsheetRequest req = new CopySheetToAnotherSpreadsheetRequest();
            req.setDestinationSpreadsheetId(spreadsheetId);
            try {
                SheetProperties resp = GoogleServiceFactory.getSheetsService().sheets().copyTo(fromSpreadsheetId, sheetId, req).execute();
                if (resp != null) {
                    sheet = GoogleDocsUtils.getSheetById(spreadsheetId, resp.getSheetId());
                }
            }
            catch (IOException e) {
                throw new GoogleException("Cannot execute batch command on %s", e, fromSpreadsheetId);
            }
        }
        return sheet == null ? null : new GoogleSheet(spreadsheetId, sheet);
    }

    /**
     * Attempts to move the spreadsheet to the folder
     *
     * @param folderId Folder to move the spreadsheet to
     * @throws GoogleException If the move failed or the folder doesn't exist
     */
    public void move(String folderId) throws GoogleException {

        // Check for rubbish
        if (folderId != null && !folderId.isEmpty()) {
            try {

                // Is this actually a folder
                GoogleFile folder = new GoogleFile(folderId);
                if (!folder.isFolder()) {
                    throw new GoogleException("Destination %s [%s] isn't a folder", folder.getName(), folderId);
                }

                // Retrieve the existing parents to remove
                GoogleFile file = new GoogleFile(spreadsheetId);
                String previousParents = file.getFile().getParents() == null ? "" : String.join(",", file.getFile().getParents());
                try {

                    // Move the file to the folder
                    GoogleServiceFactory.getFilesService().update(spreadsheetId, null)
                            .setAddParents(folderId)
                            .setRemoveParents(previousParents)
                            .setFields("id, parents")
                            .execute();
                }
                catch (IOException e) {
                    throw new GoogleException("Cannot move file %s to folder %s", e, spreadsheetId, folderId);
                }

            }
            catch (GoogleException e) {
                throw new GoogleException("Cannot open destination folder %s", folderId);
            }
        }

    }

    /**
     * Convenience method to delete the spreadsheet
     *
     * @throws GoogleException If the deletion fails
     */
    public void delete() throws GoogleException {

        // Deleting files can only be done using the Files API
        GoogleFile file = new GoogleFile(spreadsheetId);
        file.delete();
    }

    /**
     * Convenience method to rename the sheet
     *
     * @param name New name to set the spreadsheet to
     * @throws GoogleException If there is a sheet with that name
     */
    public void rename(String name) throws GoogleException {

        // Renaming files can only be done using the Files API
        GoogleFile file = new GoogleFile(spreadsheetId);
        file.renameTo(name);
    }

    /**
     * Returns the name of this spreadsheet file
     *
     * @return Name of the spreadsheet
     * @throws GoogleException If cannot locate file
     */
    public String getName() throws GoogleException {
        GoogleFile file = new GoogleFile(spreadsheetId);
        return file.getName();
    }

    /**
     * Deletes all the sheets from the spreadsheet
     *
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clear() throws GoogleException {

        // Get a list of all the sheets
        Map<String, GoogleSheet> sheets = getSheetsMap();

        // We need to batch these requests so that we don't exhaust the dreaded read request
        // rate limits
        List<Request> batchedRequestList = new LinkedList<>();

        // See if we have a Sheet1 - we can't delete all sheets, there has to be at least
        // one in a spreadsheet so rename sheet1 if it exists and add a new one
        GoogleSheet firstSheet = sheets.get("Sheet1");
        if (firstSheet != null) {
            firstSheet.rename("old_" + System.currentTimeMillis(), batchedRequestList);
        }

        // Add a default sheet
        AddSheetRequest addSheetRequest = new AddSheetRequest();
        addSheetRequest.setProperties(new SheetProperties().setTitle("Sheet1"));
        Request request = new Request();
        request.setAddSheet(addSheetRequest);
        batchedRequestList.add(request);

        // Delete all the other sheets
        for (GoogleSheet sheet : sheets.values()) {
            DeleteSheetRequest req = new DeleteSheetRequest();
            req.setSheetId(sheet.getSheetId());
            request = new Request();
            request.setDeleteSheet(req);
            batchedRequestList.add(request);
        }

        // Execute the batch
        BatchUpdateSpreadsheetResponse resp = GoogleDocsUtils.executeBatchRequest(spreadsheetId, batchedRequestList.toArray(new Request[0]));
    }

    /**
     * Gets the sheet using its name/title. NAMES are NOT stable and can change
     *
     * @param name Name of the sheet
     * @return GoogleSheet or null if not found
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet getSheetByName(String name) throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetByName(spreadsheetId, name);
        return sheet == null ? null : new GoogleSheet(spreadsheetId, sheet);
    }

    /**
     * Gets the sheet using it's unique ID. IDs are stable and never change
     *
     * @param sheetId Sheet ID
     * @return GoogleSheet if it is found
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet getSheetById(int sheetId) throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
        return sheet == null ? null : new GoogleSheet(spreadsheetId, sheet);
    }

    /**
     * Adds a new sheet to the spreadsheet at the end of the list
     * If a sheet already exists, then this is returned
     *
     * @param name     Name to give the new sheet
     * @return GoogleSheet if successful
     * @throws GoogleException If sheet cannot be created
     */
    public GoogleSheet addSheet(String name) throws GoogleException {
        return addSheet(name, 9999, false);
    }

    /**
     * Adds a new sheet to the spreadsheet at the position specified
     * If a sheet already exists and deleteIfExists is false, then this is returned
     * and if the position is out of range, it is hauled back into range
     *
     * @param name     Name to give the new sheet
     * @param position Position, starting at 0
     * @param deleteIfExists If true, will delete the sheet if it exists and create a new one
     * @return GoogleSheet if successful
     * @throws GoogleException If sheet cannot be created
     */
    public GoogleSheet addSheet(String name, int position, boolean deleteIfExists) throws GoogleException {
        GoogleSheet sheet = getSheetByName(name);
        if (sheet != null && deleteIfExists) {
            sheet.delete();
            sheet = null;
        }
        if (sheet == null) {

            // Make sure the index is in range
            int size = GoogleDocsUtils.getSpreadsheet(spreadsheetId).getSheets().size();
            position = Math.min(position, size);

            // Add the sheet
            AddSheetRequest addSheetRequest = new AddSheetRequest();
            addSheetRequest.setProperties(new SheetProperties().setTitle(name).setIndex(position));
            Request request = new Request();
            request.setAddSheet(addSheetRequest);
            BatchUpdateSpreadsheetResponse resp = GoogleDocsUtils.executeBatchRequest(spreadsheetId, request);

            // Get the new sheet
            Sheet newSheet = GoogleDocsUtils.getSheetById(spreadsheetId, resp.getReplies().get(0).getAddSheet().getProperties().getSheetId());
            if (newSheet == null) {
                throw new GoogleException("New sheet not found");
            }
            else {
                sheet = new GoogleSheet(spreadsheetId, newSheet);
            }
        }
        return sheet;
    }

    /**
     * Gets a list of the current sheets in the spreadsheet
     *
     * @return List of sheets
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public List<GoogleSheet> getSheets() throws GoogleException {
        return new ArrayList<>(getSheetsMap().values());
    }

    /**
     * Gets a map of the current sheets in the spreadsheet keyed on their names
     *
     * @return Map of sheets keyed on their names
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public Map<String, GoogleSheet> getSheetsMap() throws GoogleException {
        Map<String, GoogleSheet> ret = new LinkedHashMap<>();
        List<Sheet> sheets = GoogleDocsUtils.getSpreadsheet(spreadsheetId).getSheets();
        if (sheets != null) {
            for (Sheet sheet : sheets) {
                GoogleSheet tmp = new GoogleSheet(spreadsheetId, sheet);
                ret.put(tmp.getName(), tmp);
            }
        }
        return ret;
    }
}
