/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import lombok.ToString;
import lombok.extern.slf4j.Slf4j;

/**
 * A useful way of building and applying formatting to a GoogleSheet
 * It uses a builder pattern to help create a set of format options
 * that can be applied to multiple ranges
 */
@SuppressWarnings("unused")
@Slf4j
@ToString
public class CellFormatter {

    private final GoogleSheet sheet;
    private final String[] ranges;
    private java.awt.Color backgroundColor;
    private java.awt.Color foregroundColor;
    private Boolean bold;
    private Integer fontSize;
    private GoogleSheet.HorizontalAlignment horizontalAlignment;
    private GoogleSheet.VerticalAlignment verticalAlignment;
    private GoogleSheet.WrapStrategy wrapStrategy;
    private GoogleSheet.NumberType numberType;
    private String numberPattern;
    private Integer paddingLeft;
    private Integer paddingRight;
    private Integer paddingTop;
    private Integer paddingBottom;

    /**
     * Creates a formatter builder for the specified sheet and ranges
     *
     * @param googleSheet Sheet to format cells
     * @param ranges      Array of ranges to apply the formatting to e.g. A1,C3:D
     */
    public CellFormatter(GoogleSheet googleSheet, String[] ranges) {
        this.sheet = googleSheet;
        this.ranges = ranges;
    }

    /**
     * Sets the background color
     *
     * @param backgroundColor Color to use for background
     * @return GoogleFormater for chaining
     */
    public CellFormatter backgroundColor(java.awt.Color backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    /**
     * Sets the foreground color
     *
     * @param foregroundColor Color to use for foreground
     * @return GoogleFormater for chaining
     */
    public CellFormatter foregroundColor(java.awt.Color foregroundColor) {
        this.foregroundColor = foregroundColor;
        return this;
    }

    /**
     * Sets the font to be bold
     *
     * @return GoogleFormater for chaining
     */

    public CellFormatter bold() {
        bold = true;
        return this;
    }

    /**
     * Sets the font size
     *
     * @param fontSize in points e.g. 12
     * @return GoogleFormater for chaining
     */
    public CellFormatter fontSize(Integer fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    /**
     * Sets the horizontal alignment to left
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignLeft() {
        horizontalAlignment = GoogleSheet.HorizontalAlignment.LEFT;
        return this;
    }

    /**
     * Sets the horizontal alignment to right
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignRight() {
        horizontalAlignment = GoogleSheet.HorizontalAlignment.RIGHT;
        return this;
    }

    /**
     * Sets the horizontal alignment to center
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignCenter() {
        horizontalAlignment = GoogleSheet.HorizontalAlignment.CENTER;
        return this;
    }

    /**
     * Sets the vertical alignment to bottom
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignBottom() {
        verticalAlignment = GoogleSheet.VerticalAlignment.BOTTOM;
        return this;
    }

    /**
     * Sets the vertical alignment to top
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignTop() {
        verticalAlignment = GoogleSheet.VerticalAlignment.TOP;
        return this;
    }

    /**
     * Sets the vertical alignment to middle
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter alignMiddle() {
        verticalAlignment = GoogleSheet.VerticalAlignment.MIDDLE;
        return this;
    }

    /**
     * Sets the text wrapping to be clipped to the cell boundary
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter wrapClip() {
        wrapStrategy = GoogleSheet.WrapStrategy.CLIP;
        return this;
    }

    /**
     * Sets the text wrapping to overflow on to adjacent cells
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter wrapOverflow() {
        wrapStrategy = GoogleSheet.WrapStrategy.OVERFLOW_CELL;
        return this;
    }

    /**
     * Sets the text wrapping to fit within the cell
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter wrap() {
        wrapStrategy = GoogleSheet.WrapStrategy.WRAP;
        return this;
    }

    /**
     * Sets the display type of the cell to be text
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter text() {
        numberType = GoogleSheet.NumberType.TEXT;
        return this;
    }

    /**
     * Sets the display type of the cell to be number
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter number() {
        numberType = GoogleSheet.NumberType.NUMBER;
        return this;
    }

    /**
     * Sets the display type of the cell to be percent
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter percent() {
        numberType = GoogleSheet.NumberType.PERCENT;
        return this;
    }

    /**
     * Sets the display type of the cell to be currency
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter currency() {
        numberType = GoogleSheet.NumberType.CURRENCY;
        return this;
    }

    /**
     * Sets the display type of the cell to be date
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter date() {
        numberType = GoogleSheet.NumberType.DATE;
        return this;
    }

    /**
     * Sets the display type of the cell to be time
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter time() {
        numberType = GoogleSheet.NumberType.TIME;
        return this;
    }

    /**
     * Sets the display type of the cell to be date/time
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter dateTime() {
        numberType = GoogleSheet.NumberType.DATE_TIME;
        return this;
    }

    /**
     * Sets the display type of the cell to be scientific
     *
     * @return GoogleFormater for chaining
     */
    public CellFormatter scientific() {
        numberType = GoogleSheet.NumberType.SCIENTIFIC;
        return this;
    }

    /**
     * For numeric display types, sets the pattern to use to format the number
     *
     * @param pattern Pattern for display e.g. "mmm yyyy" for date types
     * @return GoogleFormater for chaining
     */
    public CellFormatter numberPattern(String pattern) {
        this.numberPattern = pattern;
        return this;
    }

    /**
     * Sets the left padding
     *
     * @param padding Pixels
     * @return GoogleFormater for chaining
     */
    public CellFormatter paddingLeft(int padding) {
        this.paddingLeft = padding;
        return this;
    }

    /**
     * Sets the right padding
     *
     * @param padding Pixels
     * @return GoogleFormater for chaining
     */
    public CellFormatter paddingRight(int padding) {
        this.paddingRight = padding;
        return this;
    }

    /**
     * Sets the top padding
     *
     * @param padding Pixels
     * @return GoogleFormater for chaining
     */
    public CellFormatter paddingTop(int padding) {
        this.paddingTop = padding;
        return this;
    }

    /**
     * Sets the bottom padding
     *
     * @param padding Pixels
     * @return GoogleFormater for chaining
     */
    public CellFormatter paddingBottom(int padding) {
        this.paddingBottom = padding;
        return this;
    }

    /**
     * Sets the padding
     *
     * @param padding Pixels
     * @return GoogleFormater for chaining
     */
    public CellFormatter padding(int padding) {
        paddingLeft = padding;
        paddingRight = padding;
        paddingTop = padding;
        paddingBottom = padding;
        return this;
    }

    /**
     * Applies the format to the specified ranges
     *
     * @throws GoogleException If there is a problem applying the range
     */
    public void apply() throws GoogleException {
        if (sheet != null && ranges != null) {
            for (String range : ranges) {
                if (range != null) {
                    sheet.formatCells(range, backgroundColor, foregroundColor, bold, fontSize, horizontalAlignment, verticalAlignment, wrapStrategy,
                            numberType, numberPattern, paddingLeft, paddingTop, paddingBottom, paddingRight);
                }
            }
        }
    }

}
