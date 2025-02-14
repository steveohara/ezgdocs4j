/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import com.pivotal.google.GoogleServiceFactory;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.io.StringReader;
import java.time.LocalDate;
import java.util.*;

/**
 * A metaphor for a Google sheet
 * @noinspection ALL
 */
@Slf4j
@SuppressWarnings({"unused", "ResultOfMethodCallIgnored"})
public class GoogleSheet {
    @Getter
    private String name;
    @Getter
    private final String type;
    @Getter
    private final int sheetId;
    @Getter
    private final int position;
    @Getter
    private boolean hidden;

    @Getter(AccessLevel.PROTECTED)
    private final String spreadsheetId;
    private List<Request> batchedRequestList = null;

    public enum ValueRenderOption {
        FORMATTED_VALUE,   // Values will be calculated & formatted in the response according to the cell's formatting
        UNFORMATTED_VALUE, // Values will be calculated, but not formatted
        FORMULA            // Values will not be calculated
    }

    public enum MergeType {
        MERGE_ALL,       // Merge all cells in the range
        MERGE_COLUMNS,   // Merge just the columns
        MERGE_ROWS       // Merge just the rows
    }

    public enum VerticalAlignment {
        TOP, MIDDLE, BOTTOM
    }

    public enum HorizontalAlignment {
        LEFT, CENTER, RIGHT
    }

    public enum WrapStrategy {
        WRAP,          // Wrap lines that are longer than the cell width onto a new line. Single words that are longer than a line are wrapped at the character level.
        OVERFLOW_CELL, // Overflow lines into the next cell, so long as that cell is empty. If the next cell over is non-empty, this behaves the same as CLIP.
        CLIP           // Clip lines that are longer than the cell width.
    }

    public enum NumberType {
        TEXT,      // Text formatting, e.g 1000.12
        NUMBER,    // Number formatting, e.g, 1,000.12
        PERCENT,   // Percent formatting, e.g 10.12%
        CURRENCY,  // Currency formatting, e.g $1,000.12
        DATE,      // Date formatting, e.g 9/26/2008
        TIME,      // Time formatting, e.g 3:59:00 PM
        DATE_TIME, // Date+Time formatting, e.g 9/26/08 15:59:00
        SCIENTIFIC // Scientific number formatting, e.g 1.01E+03
    }

    /**
     * Creates a GoogleSheet object belonging to the specified spreadsheet
     *
     * @param spreadsheetId Spreadsheet ID
     * @param sheet         The Sheet object to get data from
     */
    protected GoogleSheet(String spreadsheetId, Sheet sheet) {
        this.spreadsheetId = spreadsheetId;
        SheetProperties props = sheet.getProperties();
        this.name = props.getTitle();
        this.sheetId = props.getSheetId();
        this.position = props.getIndex();
        this.type = props.getSheetType();
        this.hidden = props.getHidden() != null && props.getHidden();
    }

    /**
     * Deletes the sheet if it exists, otherwise
     *
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void delete() throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
        if (sheet != null) {
            DeleteSheetRequest req = new DeleteSheetRequest();
            req.setSheetId(sheetId);
            Request request = new Request();
            request.setDeleteSheet(req);
            executeBatchRequest(request);
        }
    }

    /**
     * Copies the sheet if it exists
     * If the duplicate already exists, returns the existing sheet
     *
     * @param name Name to give the duplicate sheet
     * @return The duplicated sheet or null if the command is batched
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet duplicate(String name) throws GoogleException {

        // Check if the sheet already exists
        Sheet sheet = GoogleDocsUtils.getSheetByName(spreadsheetId, name);
        if (sheet == null) {
            DuplicateSheetRequest req = new DuplicateSheetRequest();
            req.setNewSheetName(name);
            Request request = new Request();
            request.setDuplicateSheet(req);
            BatchUpdateSpreadsheetResponse resp = executeBatchRequest(request);

            // Get the new sheet
            if (resp != null) {
                List<Response> replies = resp.getReplies();
                if (replies != null && !replies.isEmpty()) {
                    sheet = GoogleDocsUtils.getSheetById(spreadsheetId, replies.get(0).getDuplicateSheet().getProperties().getSheetId());
                }
            }
            if (sheet == null && batchedRequestList != null) {
                throw new GoogleException("New sheet not found");
            }
        }
        return sheet == null ? null : new GoogleSheet(spreadsheetId, sheet);
    }

    /**
     * Copies the sheet if it exists from the current spreadsheet top the target
     * If the duplicate already exists, returns the existing sheet
     *
     * @param targetSpreadsheetId The spreadsheet to copy the sheet to
     * @return The duplicated sheet or null if the command is batched
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public GoogleSheet copyTo(String targetSpreadsheetId) throws GoogleException {

        // Check if the sheet already exists
        Sheet sheet = GoogleDocsUtils.getSheetByName(targetSpreadsheetId, name);
        if (sheet == null) {
            CopySheetToAnotherSpreadsheetRequest req = new CopySheetToAnotherSpreadsheetRequest();
            req.setDestinationSpreadsheetId(targetSpreadsheetId);
            try {
                SheetProperties resp = GoogleServiceFactory.getSheetsService().sheets().copyTo(spreadsheetId, sheetId, req).execute();
                if (resp != null) {
                    sheet = GoogleDocsUtils.getSheetById(targetSpreadsheetId, resp.getSheetId());
                }
            }
            catch (IOException e) {
                throw new GoogleException("Cannot execute batch command on %s", e, spreadsheetId);
            }
        }
        return sheet == null ? null : new GoogleSheet(targetSpreadsheetId, sheet);
    }

    /**
     * Returns the data for the specified range in A1 notation
     * The data returned is either the formulas and raw data or the unformatted display
     * values depending on whether formulas is true or not
     *
     * @param range        Range e.g. A1:C6
     * @param renderOption How the returned data should be formatted
     * @return List of rows and columns
     * @throws GoogleException If it cannot get the data or the range is not correctly specified
     */
    public List<List<Object>> getData(String range, ValueRenderOption renderOption) throws GoogleException {
        Sheets.Spreadsheets.Values valuesService = GoogleServiceFactory.getSheetsService().values();
        String canonicalRange = name + (range == null ? "" : ("!" + range));
        try {
            ValueRange values = valuesService.get(spreadsheetId, canonicalRange)
                    .setValueRenderOption(renderOption.toString())
                    .setDateTimeRenderOption(renderOption.equals(ValueRenderOption.FORMATTED_VALUE) ? "FORMATTED_STRING" : "SERIAL_NUMBER")
                    .setMajorDimension("ROWS")
                    .execute();
            return values.getValues();
        }
        catch (IOException e) {
            throw new GoogleException("Cannot get sheet data for %s", e, canonicalRange);
        }
    }

    /**
     * Hides/unhides the sheet
     * Doesn't do anything if the sheet doesn't exist
     *
     * @param hidden True if the sheet should be hidden, false for visible
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setHidden(boolean hidden) throws GoogleException {
        SheetProperties props = new SheetProperties();
        props.setSheetId(sheetId);
        props.setHidden(hidden);
        UpdateSheetPropertiesRequest req = new UpdateSheetPropertiesRequest();
        req.setProperties(props);
        req.setFields("hidden");
        Request request = new Request();
        request.setUpdateSheetProperties(req);
        executeBatchRequest(request);
        this.hidden = hidden;
    }

    /**
     * Renames the sheet to a new name
     * If the new name exists, the operation cancels
     *
     * @param name Name to set the sheet to
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void rename(String name) throws GoogleException {
        rename(name, null);
    }

    /**
     * Renames the sheet to a new name
     * If the new name exists, the operation cancels
     *
     * @param name        Name to set the sheet to
     * @param requestList Optional list to add the request to
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    protected void rename(String name, List<Request> requestList) throws GoogleException {

        // Check if the sheet already exists
        Sheet sheet = GoogleDocsUtils.getSheetByName(spreadsheetId, name);
        if (sheet == null) {

            // Get the values for current sheet
            sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
            if (sheet != null) {
                SheetProperties props = sheet.getProperties();
                props.setTitle(name);
                UpdateSheetPropertiesRequest req = new UpdateSheetPropertiesRequest();
                req.setProperties(props);
                req.setFields("title");
                Request request = new Request();
                request.setUpdateSheetProperties(req);

                // Check to see how to execute it
                if (requestList != null) {
                    requestList.add(request);
                }
                else {
                    executeBatchRequest(request);
                }

                // Taking a bit of a chance here, because the request may not have been
                // executed so the name may not yet have changed on the server
                this.name = name;
            }
        }
    }

    /**
     * Clears all the cell contents and formatting information
     *
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clear() throws GoogleException {
        clear(null);
    }

    /**
     * Clears all the cell contents and formatting information
     *
     * @param range Range in R1C1 notation to clear or null for whole sheet
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clear(String range) throws GoogleException {
        UpdateCellsRequest req = new UpdateCellsRequest();
        GridRange gridRange = range == null ? new GridRange() : GoogleDocsUtils.getGridRange(sheetId, range);
        gridRange.setSheetId(sheetId);
        req.setRange(gridRange);
        req.setFields("*");
        Request request = new Request();
        request.setUpdateCells(req);
        executeBatchRequest(request);
    }

    /**
     * Clears the values in the given sheet, but leaves all the formatting behind
     *
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clearData() throws GoogleException {
        clearData(null);
    }

    /**
     * Clears the values in the given sheet, but leaves all the formatting behind
     *
     * @param range Range to clear in A:B notation
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clearData(String range) throws GoogleException {
        Sheets.Spreadsheets.Values valuesService = GoogleServiceFactory.getSheetsService().values();
        try {
            valuesService.clear(spreadsheetId, name + (range == null ? "" : ("!" + range)), new ClearValuesRequest()).execute();
        }
        catch (IOException e) {
            throw new GoogleException("Cannot clear sheet data for %s", e, name);
        }
    }

    /**
     * Sets or removes a note from a cell or range of cells
     *
     * @param range R1C1 Notation for the range
     * @param note  Note to add to a column - null, removes the notes
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setNote(String range, String note) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setNote(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), note);
    }

    /**
     * Sets or removes a note from a cell or range of cells
     *
     * @param startRow    Start row of range to add/remove note
     * @param startColumn Start column of range to add/remove note
     * @param endRow      End row of range to add/remove note (must be greater than startRow)
     * @param endColumn   End column of range to add/remove note (must be greater than endColumn)
     * @param note        Note to add to a column - null, removes the notes
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setNote(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, String note) throws GoogleException {

        // Get a repeating cell request
        RepeatCellRequest req = GoogleDocsUtils.getRepeatCellRequest(sheetId, startRow, startColumn, endRow, endColumn);
        CellData cellData = new CellData();
        cellData.setNote(note);
        req.setCell(cellData);
        req.setFields("note");

        // Execute the request
        Request request = new Request();
        request.setRepeatCell(req);
        executeBatchRequest(request);
    }

    /**
     * Appends text formatted as lines of CSV
     * If rangeStart is null, then the data is appended to the end of the
     * current data range
     *
     * @param rangeStart    Where to start to append the data from R1C1 notation
     * @param csvText       CSV data, line breaks
     * @param includeHeader True if the header should be included
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void appendCsvValues(String rangeStart, String csvText, boolean includeHeader) throws GoogleException {
        try (CSVReader reader = new CSVReader(new StringReader(csvText))) {

            // Buffer for the sheet data
            List<List<Object>> rows = new ArrayList<>();
            String[] values;
            boolean firstRow = true;
            while ((values = reader.readNext()) != null) {
                if (!firstRow || includeHeader) {
                    rows.add(Arrays.asList(values));
                }
                firstRow = false;
            }
            appendValues(rangeStart, rows);
        }
        catch (IOException | CsvValidationException e) {
            throw new GoogleException("Cannot read CSV data", e);
        }
    }

    /**
     * Appends the Object values (single row) to the sheet
     * If rangeStart is null, then the data is appended to the end of the
     * current data range
     *
     * @param rangeStart Where to start to append the data from R1C1 notation
     * @param values     Array of Objects
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void appendValues(String rangeStart, Object... values) throws GoogleException {
        List<List<Object>> rows = new ArrayList<>();
        rows.add(Arrays.asList(values));
        appendValues(rangeStart, rows);
    }

    /**
     * Appends the List of Lists of Object values to the sheet
     * If rangeStart is null, then the data is appended to the end of the
     * current data range
     *
     * @param rangeStart Where to start to append the data from R1C1 notation
     * @param values     List of List of Objects
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void appendValues(String rangeStart, List<List<Object>> values) throws GoogleException {
        Sheets.Spreadsheets.Values valuesService = GoogleServiceFactory.getSheetsService().values();
        try {
            Sheets.Spreadsheets.Values.Append append = valuesService.append(spreadsheetId, name + (rangeStart == null ? "" : ("!" + rangeStart)), new ValueRange().setValues(values));
            append.setValueInputOption("USER_ENTERED");
            append.setIncludeValuesInResponse(false);
            append.execute();
        }
        catch (IOException e) {
            throw new GoogleException("Cannot append data to %s at location %s", e, name, rangeStart);
        }
    }

    /**
     * Sets the number of rows to freeze
     *
     * @param rows    Rows to freeze (null means leave)
     * @param columns Columns to freeze (null means leave)
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void freezeRowsAndColumns(Integer rows, Integer columns) throws GoogleException {

        // If we have something to do
        if (rows != null || columns != null) {
            List<String> fields = new ArrayList<>();
            UpdateSheetPropertiesRequest req = new UpdateSheetPropertiesRequest();
            SheetProperties props = new SheetProperties();
            props.setSheetId(sheetId);

            // Specify the grid affected
            GridProperties gridProps = new GridProperties();
            if (rows != null) {
                gridProps.setFrozenRowCount(rows);
                fields.add("gridProperties.frozenRowCount");
            }
            if (columns != null) {
                gridProps.setFrozenColumnCount(columns);
                fields.add("gridProperties.frozenColumnCount");
            }

            // Set the fields to be used to be updated
            req.setFields(String.join(",", fields));
            props.setGridProperties(gridProps);
            req.setProperties(props);

            // Build and execute the request
            Request request = new Request();
            request.setUpdateSheetProperties(req);
            executeBatchRequest(request);
        }
    }

    /**
     * Returns a formatter to use instead of the more verbose null based parameter
     * lists of the other forms of this method.
     * The formatter is a builder pattern that can apply the format to multiple cell
     * ranges and provides a more syntactically succinct method of applying formatting
     *
     * @param ranges Array of ranges to apply the formatting to
     * @return A GoogleFormatter instance
     */
    public CellFormatter formatCells(String... ranges) {
        return new CellFormatter(this, ranges);
    }

    /**
     * Formats the contents of the cells
     * If any argument is null, then it is ignored.
     *
     * @param range               R1C1 Notation for the range
     * @param backgroundColor     Background color to apply
     * @param foregroundColor     Foreground color of the text
     * @param bold                Set to bold
     * @param fontSize            Font size rto apply
     * @param horizontalAlignment Horizontal alignment within each cell
     * @param verticalAlignment   Vertical alignment within each cell
     * @param wrapStrategy        Type of cell content wrapping
     * @param numberType          Type of the number
     * @param numberPattern       Pattern to use to show the number
     * @param paddingLeft         Padding in pixels for left of cell
     * @param paddingTop          Padding in pixels for top of cell
     * @param paddingBottom       Padding in pixels for bottom of cell
     * @param paddingRight        Padding in pixels for right of cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void formatCells(String range, java.awt.Color backgroundColor, java.awt.Color foregroundColor, Boolean bold, Integer fontSize, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, WrapStrategy wrapStrategy,
                            NumberType numberType, String numberPattern, Integer paddingLeft, Integer paddingTop, Integer paddingBottom, Integer paddingRight) throws GoogleException {

        // Convert the range to its numeric form
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        formatCells(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(),
                backgroundColor, foregroundColor, bold, fontSize, horizontalAlignment, verticalAlignment, wrapStrategy, numberType, numberPattern,
                paddingLeft, paddingTop, paddingBottom, paddingRight);
    }

    /**
     * Formats the contents of the cells
     * If any argument is null, then it is ignored which is also true of the range itself.
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow            Start row of range to format
     * @param startColumn         Start column of range to format
     * @param endRow              End row of range to format (must be greater than startRow)
     * @param endColumn           End column of range to format (must be greater than endColumn)
     * @param backgroundColor     Background color to apply
     * @param foregroundColor     Foreground color of the text
     * @param bold                Set to bold
     * @param fontSize            Font size rto apply
     * @param horizontalAlignment Horizontal alignment within each cell
     * @param verticalAlignment   Vertical alignment within each cell
     * @param wrapStrategy        Type of cell content wrapping
     * @param numberType          Type of the number
     * @param numberPattern       Pattern to use to show the number
     * @param paddingLeft         Padding in pixels for left of cell
     * @param paddingTop          Padding in pixels for top of cell
     * @param paddingBottom       Padding in pixels for bottom of cell
     * @param paddingRight        Padding in pixels for right of cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void formatCells(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, java.awt.Color backgroundColor, java.awt.Color foregroundColor, Boolean bold, Integer fontSize, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, WrapStrategy wrapStrategy,
                            NumberType numberType, String numberPattern, Integer paddingLeft, Integer paddingTop, Integer paddingBottom, Integer paddingRight) throws GoogleException {

        // Get a repeating cell request
        RepeatCellRequest req = GoogleDocsUtils.getRepeatCellRequest(sheetId, startRow, startColumn, endRow, endColumn);

        // Formatting rules
        TextFormat textFormat = new TextFormat();
        CellData cell = new CellData();
        req.setCell(cell);

        // Define the cell formatting
        Set<String> fields = new HashSet<>();
        CellFormat cellFormat = new CellFormat();
        if (backgroundColor != null) {
            cellFormat.setBackgroundColor(GoogleDocsUtils.convertColorToGoogle(backgroundColor));
            fields.add("userEnteredFormat.backgroundColor");
        }
        if (horizontalAlignment != null) {
            cellFormat.setHorizontalAlignment(horizontalAlignment.toString());
            fields.add("userEnteredFormat.horizontalAlignment");
        }
        if (verticalAlignment != null) {
            cellFormat.setVerticalAlignment(verticalAlignment.toString());
            fields.add("userEnteredFormat.verticalAlignment");
        }
        if (wrapStrategy != null) {
            cellFormat.setWrapStrategy(wrapStrategy.toString());
            fields.add("userEnteredFormat.wrapStrategy");
        }
        Padding padding = new Padding();
        if (paddingLeft != null) {
            padding.setLeft(paddingLeft);
        }
        if (paddingTop != null) {
            padding.setTop(paddingTop);
        }
        if (paddingBottom != null) {
            padding.setBottom(paddingBottom);
        }
        if (paddingRight != null) {
            padding.setRight(paddingRight);
        }
        if (paddingLeft != null || paddingTop != null || paddingBottom != null || paddingRight != null) {
            cellFormat.setPadding(padding);
            fields.add("userEnteredFormat.padding");
        }
        if (numberType != null) {
            NumberFormat numberFormat = new NumberFormat();
            numberFormat.setType(numberType.toString());
            fields.add("userEnteredFormat.numberFormat");
            if (numberPattern != null) {
                numberFormat.setPattern(numberPattern);
            }
            cellFormat.setNumberFormat(numberFormat);
        }
        cell.setUserEnteredFormat(cellFormat);

        // Text formatting
        if (foregroundColor != null) {
            textFormat.setForegroundColor(GoogleDocsUtils.convertColorToGoogle(foregroundColor));
            fields.add("userEnteredFormat.textFormat");
        }
        if (bold != null) {
            textFormat.setBold(bold);
            fields.add("userEnteredFormat.textFormat");
        }
        if (fontSize != null) {
            textFormat.setFontSize(fontSize);
            fields.add("userEnteredFormat.textFormat");
        }
        cellFormat.setTextFormat(textFormat);

        // Update all fields
        if (!fields.isEmpty()) {
            req.setFields(String.join(",", fields));
        }

        // Execute the request
        Request request = new Request();
        request.setRepeatCell(req);
        executeBatchRequest(request);
    }

    /**
     * Adds a filter view to the sheet
     *
     * @param range    R1C1 Notation for the range
     * @param name     Name of the filter view to create
     * @param criteria Map of filter criteria keyed on the column index e.g. "0"
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void addFilter(String range, String name, Map<String, FilterCriteria> criteria) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        addFilter(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), name, criteria);
    }

    /**
     * Adds a filter view to the sheet
     *
     * @param startRow    Start row of range of data to be filtered
     * @param startColumn Start column of range data to be filtered
     * @param endRow      End row of range data to be filtered (must be greater than startRow)
     * @param endColumn   End column of range data to be filtered (must be greater than endColumn)
     * @param name        Name of the filter view to create
     * @param criteria    Map of filter criteria keyed on the column index e.g. "0"
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void addFilter(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, String name, Map<String, FilterCriteria> criteria) throws GoogleException {

        // Create the request
        AddFilterViewRequest req = new AddFilterViewRequest();
        FilterView filterView = new FilterView();
        filterView.setTitle(name);
        filterView.setRange(GoogleDocsUtils.getGridRange(sheetId, startRow, startColumn, endRow, endColumn));
        filterView.setCriteria(criteria);
        req.setFilter(filterView);

        // Execute the request
        Request request = new Request();
        request.setAddFilterView(req);
        executeBatchRequest(request);
    }

    /**
     * Gets a filter using its name
     *
     * @param name Name of the filter view
     * @return FilterView or null if not found
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public FilterView getFilterByName(String name) throws GoogleException {
        return getFilterMap().get(name);
    }

    /**
     * Gets a map of all filter keyed on their names
     *
     * @return Returns a map of filter views keyed on the name
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public Map<String, FilterView> getFilterMap() throws GoogleException {
        Map<String, FilterView> ret = new TreeMap<>(String.CASE_INSENSITIVE_ORDER);
        Sheet sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
        if (sheet != null) {
            List<FilterView> views = sheet.getFilterViews();
            if (views != null) {
                for (FilterView view : views) {
                    ret.put(view.getTitle(), view);
                }
            }
        }
        return ret;
    }

    /**
     * Deletes a filter using its name
     *
     * @param name Name of the filter view
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void deleteFilterByName(String name) throws GoogleException {

        // Check we have a filter of that name
        FilterView view = getFilterByName(name);
        if (view != null) {

            // Create the request
            DeleteFilterViewRequest req = new DeleteFilterViewRequest();
            req.setFilterId(view.getFilterViewId());

            // Execute the request
            Request request = new Request();
            request.setDeleteFilterView(req);
            executeBatchRequest(request);
        }
    }

    /**
     * Auto-resizes the specified column ranges
     *
     * @param ranges R1C1 Notation for the ranges
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void autoResizeColumns(String... ranges) throws GoogleException {
        for (String range : ranges) {
            GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
            autoResizeColumns(gridRange.getStartColumnIndex(), gridRange.getEndColumnIndex());
        }
    }

    /**
     * Auto-resizes the specified column range
     *
     * @param startColumn Start column of range to format
     * @param endColumn   End column of range to format (must be greater than startColumn)
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void autoResizeColumns(int startColumn, int endColumn) throws GoogleException {

        // Create the request
        AutoResizeDimensionsRequest req = new AutoResizeDimensionsRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setStartIndex(startColumn);
        range.setEndIndex(endColumn);
        range.setDimension("COLUMNS");
        req.setDimensions(range);

        // Execute the request
        Request request = new Request();
        request.setAutoResizeDimensions(req);
        executeBatchRequest(request, true);
    }

    /**
     * Sets the width of the column(s) in pixels
     *
     * @param width   Width of the column in pixels
     * @param columns Columns in R1C1 format to set the size of e.g. 'C'
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setColumnWidth(int width, String... columns) throws GoogleException {
        for (String range : columns) {
            GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
            setColumnWidth(width, gridRange.getStartColumnIndex(), gridRange.getEndColumnIndex());
        }
    }

    /**
     * Sets the width of the column in pixels
     *
     * @param width       Width of the column in pixels
     * @param startColumn Start column of range to change size
     * @param endColumn   End column of range to change size
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setColumnWidth(int width, int startColumn, Integer endColumn) throws GoogleException {

        // Check for idiocy
        if (width < 0) {
            throw new GoogleException("Invalid column width: " + width);
        }
        if (startColumn < 0) {
            throw new GoogleException("Invalid start column number: " + startColumn);
        }
        if (endColumn == null) {
            endColumn = startColumn + 1;
        }
        else if (endColumn <= startColumn) {
            throw new GoogleException("Invalid end column number: " + startColumn);
        }

        // Create the request
        UpdateDimensionPropertiesRequest req = new UpdateDimensionPropertiesRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setStartIndex(startColumn);
        range.setEndIndex(endColumn);
        range.setDimension("COLUMNS");
        req.setRange(range);
        DimensionProperties props = new DimensionProperties();
        props.setPixelSize(width);
        req.setProperties(props);
        req.setFields("pixelSize");

        // Execute the request
        Request request = new Request();
        request.setUpdateDimensionProperties(req);
        executeBatchRequest(request, true);
    }

    /**
     * Sets the height of the row(s) in pixels
     *
     * @param height Height of the row in pixels
     * @param rows   Rows in R1C1 format to set the size of e.g. 'C1' == Row 1
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setRowHeight(int height, String... rows) throws GoogleException {
        for (String range : rows) {
            GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
            setRowHeight(height, gridRange.getStartRowIndex(), gridRange.getEndRowIndex());
        }
    }

    /**
     * Sets the height of the rows in pixels
     *
     * @param height   Height of the row in pixels
     * @param startRow Start row of range to change size
     * @param endRow   End row of range to change size
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setRowHeight(int height, int startRow, Integer endRow) throws GoogleException {

        // Check for idiocy
        if (height < 0) {
            throw new GoogleException("Invalid column width: " + height);
        }
        if (startRow < 0) {
            throw new GoogleException("Invalid start column number: " + startRow);
        }
        if (endRow == null) {
            endRow = startRow + 1;
        }
        else if (endRow <= startRow) {
            throw new GoogleException("Invalid end column number: " + startRow);
        }

        // Create the request
        UpdateDimensionPropertiesRequest req = new UpdateDimensionPropertiesRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setStartIndex(startRow);
        range.setEndIndex(endRow);
        range.setDimension("ROWS");
        req.setRange(range);
        DimensionProperties props = new DimensionProperties();
        props.setPixelSize(height);
        req.setProperties(props);
        req.setFields("pixelSize");

        // Execute the request
        Request request = new Request();
        request.setUpdateDimensionProperties(req);
        executeBatchRequest(request, true);
    }

    /**
     * Applies a formula to the specified range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param range     R1C1 Notation for the range
     * @param mergeType Type of cell merge to do e.g. MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void mergeCells(String range, MergeType mergeType) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        mergeCells(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), mergeType);
    }

    /**
     * Applies a formula to the specified range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param mergeType   Type of cell merge to do e.g. MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void mergeCells(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, MergeType mergeType) throws GoogleException {

        // Create a request for the sheet
        MergeCellsRequest req = new MergeCellsRequest();
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, startRow, startColumn, endRow, endColumn);
        req.setRange(gridRange);
        req.setMergeType(mergeType.toString());

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setMergeCells(req);
        executeBatchRequest(request);
    }

    /**
     * Appends empty rows to the end of the sheet
     *
     * @param rows Number of empty rows to append
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void appendRows(int rows) throws GoogleException {

        // Create a request for the sheet
        AppendDimensionRequest req = new AppendDimensionRequest();
        req.setDimension("ROWS");
        req.setLength(rows);
        req.setSheetId(sheetId);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setAppendDimension(req);
        executeBatchRequest(request);
    }

    /**
     * Appends empty columns to the end of the sheet
     *
     * @param columns Number of empty columns to append
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void appendColumns(int columns) throws GoogleException {

        // Create a request for the sheet
        AppendDimensionRequest req = new AppendDimensionRequest();
        req.setDimension("COLUMNS");
        req.setLength(columns);
        req.setSheetId(sheetId);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setAppendDimension(req);
        executeBatchRequest(request);
    }

    /**
     * Delete rows starting at startRow all the way to the end
     *
     * @param startRow First row to delete
     * @throws GoogleException If the range is badly formed
     */
    public void deleteRows(int startRow) throws GoogleException {
        int rows = getRowCount();
        if (startRow < rows) {
            deleteRows(startRow, rows - startRow);
        }
    }

    /**
     * Delete rows starting at startRow
     *
     * @param startRow First row to delete
     * @param count    Number of rows to delete
     * @throws GoogleException If the range is badly formed
     */
    public void deleteRows(int startRow, int count) throws GoogleException {

        // Check for some oddities
        if (startRow < 0 || count <= 0) {
            return;
        }

        // Create a request for the sheet
        DeleteDimensionRequest req = new DeleteDimensionRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setDimension("ROWS");
        range.setStartIndex(startRow);
        range.setEndIndex(startRow + count);
        req.setRange(range);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setDeleteDimension(req);
        executeBatchRequest(request, true);
    }

    /**
     * Delete columns starting at startRow
     *
     * @param range R1C1 Notation for the range
     * @throws GoogleException If the range is badly formed or isn't valid
     */
    public void deleteColumns(String range) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        deleteColumns(gridRange.getStartColumnIndex(), gridRange.getEndColumnIndex() - gridRange.getStartColumnIndex());
    }

    /**
     * Delete columns starting at startColumn
     *
     * @param startColumn First column to delete
     * @param count       Number of columns to delete
     * @throws GoogleException If the range is badly formed
     */
    public void deleteColumns(int startColumn, int count) throws GoogleException {

        // Check for some oddities
        if (startColumn < 0 || count <= 0) {
            return;
        }

        // Create a request for the sheet
        DeleteDimensionRequest req = new DeleteDimensionRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setDimension("COLUMNS");
        range.setStartIndex(startColumn);
        range.setEndIndex(startColumn + count);
        req.setRange(range);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setDeleteDimension(req);
        executeBatchRequest(request, true);
    }

    /**
     * Insert rows starting at startRow
     *
     * @param startRow First row to add after
     * @param count    Number of rows to add
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void insertRows(int startRow, int count) throws GoogleException {

        // Create a request for the sheet
        InsertDimensionRequest req = new InsertDimensionRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setDimension("ROWS");
        range.setStartIndex(startRow);
        range.setEndIndex(startRow + count);
        req.setRange(range);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setInsertDimension(req);
        executeBatchRequest(request);
    }

    /**
     * Insert rows starting at startRow
     *
     * @param startColumn First column to add after
     * @param count       Number of columns to add
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void insertColumns(int startColumn, int count) throws GoogleException {

        // Create a request for the sheet
        InsertDimensionRequest req = new InsertDimensionRequest();
        DimensionRange range = new DimensionRange();
        range.setSheetId(sheetId);
        range.setDimension("COLUMNS");
        range.setStartIndex(startColumn);
        range.setEndIndex(startColumn + count);
        req.setRange(range);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setInsertDimension(req);
        executeBatchRequest(request);
    }

    /**
     * Returns the number of rows in the sheet
     *
     * @return The number of rows in the sheet
     * @throws GoogleException If the sheet cannot be opened
     */
    public int getRowCount() throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
        if (sheet != null) {
            return sheet.getProperties().getGridProperties().getRowCount();
        }
        return 0;
    }

    /**
     * Returns the number of columns in the sheet
     *
     * @return Number of columns in the sheet
     * @throws GoogleException If the sheet cannot be opened
     */
    public int getColumnCount() throws GoogleException {
        Sheet sheet = GoogleDocsUtils.getSheetById(spreadsheetId, sheetId);
        if (sheet != null) {
            return sheet.getProperties().getGridProperties().getColumnCount();
        }
        return 0;
    }

    /**
     * Clears the default (called the Basic) filter for the sheet
     *
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void clearBasicFilter() throws GoogleException {

        // Create a request for the sheet
        ClearBasicFilterRequest req = new ClearBasicFilterRequest();
        req.setSheetId(sheetId);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setClearBasicFilter(req);
        executeBatchRequest(request);
    }

    /**
     * Sets the default (called the Basic) filter for the sheet
     *
     * @param range R1C1 Notation for the range
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setBasicFilter(String range) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setBasicFilter(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex());
    }

    /**
     * Sets the default (called the Basic) filter for the sheet
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setBasicFilter(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn) throws GoogleException {

        // Clear the filer if it exists
        clearBasicFilter();

        // Create a request for the sheet
        SetBasicFilterRequest req = new SetBasicFilterRequest();
        BasicFilter filter = new BasicFilter();
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, startRow, startColumn, endRow, endColumn);
        filter.setRange(gridRange);
        req.setFilter(filter);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setSetBasicFilter(req);
        executeBatchRequest(request);
    }

    /**
     * Applies a formula to the specified range
     *
     * @param range   R1C1 Notation for the range
     * @param formula Formula to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setFormula(String range, String formula) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setFormula(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), formula);
    }

    /**
     * Applies a formula to the specified range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param formula     Formula to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setFormula(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, String formula) throws GoogleException {
        executeDataRequest(sheetId, startRow, startColumn, endRow, endColumn, new ExtendedValue().setFormulaValue(formula));
    }

    /**
     * Sets the value for the give range
     *
     * @param range R1C1 Notation for the range
     * @param value Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setValue(String range, double value) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setValue(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), value);
    }

    /**
     * Sets the value for the give range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param value       Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setValue(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, double value) throws GoogleException {
        executeDataRequest(sheetId, startRow, startColumn, endRow, endColumn, new ExtendedValue().setNumberValue(value));
    }

    /**
     * Sets the text for the given range
     *
     * @param range R1C1 Notation for the range
     * @param value Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setText(String range, String value) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setText(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), value);
    }

    /**
     * Sets the text for the given range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param value       Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setText(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, String value) throws GoogleException {
        executeDataRequest(sheetId, startRow, startColumn, endRow, endColumn, new ExtendedValue().setStringValue(value));
    }

    /**
     * Sets the text for the given range
     *
     * @param range R1C1 Notation for the range
     * @param value Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setBoolean(String range, boolean value) throws GoogleException {
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setBoolean(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), value);
    }

    /**
     * Sets the text for the given range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param value       Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setBoolean(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, boolean value) throws GoogleException {
        executeDataRequest(sheetId, startRow, startColumn, endRow, endColumn, new ExtendedValue().setBoolValue(value));
    }

    /**
     * Sets the text for the given range
     *
     * @param range R1C1 Notation for the range
     * @param date  Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setDate(String range, LocalDate date) throws GoogleException {
        double value = GoogleDocsUtils.getGoogleDateValue(date);
        GridRange gridRange = GoogleDocsUtils.getGridRange(sheetId, range);
        setDate(gridRange.getStartRowIndex(), gridRange.getStartColumnIndex(), gridRange.getEndRowIndex(), gridRange.getEndColumnIndex(), date);
    }

    /**
     * Sets the text for the given range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param startRow    Start row of range to format
     * @param startColumn Start column of range to format
     * @param endRow      End row of range to format
     * @param endColumn   End column of range to format
     * @param date        Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    public void setDate(Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, LocalDate date) throws GoogleException {
        double value = GoogleDocsUtils.getGoogleDateValue(date);
        executeDataRequest(sheetId, startRow, startColumn, endRow, endColumn, new ExtendedValue().setNumberValue(value));
    }

    /**
     * Starts a batch so that all subsequent batch updates are captured
     */
    public void batchStart() {
        batchedRequestList = new ArrayList<>();
    }

    /**
     * Clears the current batch list of requests pending execution
     */
    public void batchClear() {
        batchedRequestList = null;
    }

    /**
     * If we are in batch mode (there are some pending requests) executes
     * them as a single transaction
     *
     * @return BatchUpdateSpreadsheetResponse Response from server or null if no requests
     * @throws GoogleException If there was a problem with batch requests
     */
    public BatchUpdateSpreadsheetResponse batchExecute() throws GoogleException {
        if (batchedRequestList != null && !batchedRequestList.isEmpty()) {
            BatchUpdateSpreadsheetResponse resp = GoogleDocsUtils.executeBatchRequest(spreadsheetId, batchedRequestList.toArray(new Request[0]));
            batchedRequestList = null;
            return resp;
        }
        return null;
    }

    /**
     * Takes the requests and executes them as a batch or adds them to the
     * batch we are storing up
     *
     * @param request Request to execute or add to batch
     * @return BatchUpdateSpreadsheetResponse Response to get returned values or null if using batch mode
     * @throws GoogleException If the batch fails
     */
    private BatchUpdateSpreadsheetResponse executeBatchRequest(Request request) throws GoogleException {
        return executeBatchRequest(request, false);
    }

    /**
     * Takes the requests and executes them as a batch or adds them to the
     * batch we are storing up
     *
     * @param request         Request to execute or add to batch
     * @param absorbException True if any exception should be ignored
     * @return BatchUpdateSpreadsheetResponse Response to get returned values or null if using batch mode
     * @throws GoogleException If the batch fails
     */
    private BatchUpdateSpreadsheetResponse executeBatchRequest(Request request, boolean absorbException) throws GoogleException {
        if (batchedRequestList != null) {
            batchedRequestList.add(request);
        }
        else {
            try {
                return GoogleDocsUtils.executeBatchRequest(spreadsheetId, request);
            }
            catch (GoogleException e) {
                if (!absorbException) {
                    throw e;
                }
                else {
                    log.warn("Failed to execute last request - {}", e.getMessage());
                }
            }
        }
        return null;
    }

    /**
     * Executes a data request on the specified range
     * If only a row is specified then the whole row is updated, only a column then
     * the whole column etc.
     *
     * @param sheetId       Sheet ID to assign
     * @param startRow      Start row of range to format
     * @param startColumn   Start column of range to format
     * @param endRow        End row of range to format
     * @param endColumn     End column of range to format
     * @param extendedValue Value to apply to the cell
     * @throws GoogleException If the spreadsheet cannot be opened/found
     */
    private void executeDataRequest(int sheetId, Integer startRow, Integer startColumn, Integer endRow, Integer endColumn, ExtendedValue extendedValue) throws GoogleException {

        // Create a request for the sheet
        RepeatCellRequest req = GoogleDocsUtils.getRepeatCellRequest(sheetId, startRow, startColumn, endRow, endColumn);

        // Apply the formula
        CellData cell = new CellData();
        cell.setUserEnteredValue(extendedValue);
        req.setFields("userEnteredValue");
        req.setCell(cell);

        // Execute the request on the spreadsheet
        Request request = new Request();
        request.setRepeatCell(req);
        executeBatchRequest(request);
    }

}
