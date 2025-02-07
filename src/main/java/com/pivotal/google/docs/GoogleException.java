/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import lombok.extern.slf4j.Slf4j;

/**
 * A useful exception to delineate from system exceptions
 */
@Slf4j
public class GoogleException extends Exception {

    /**
     * Constructs an {@code GoogleException} with {@code null}
     * as its error detail message.
     */
    public GoogleException() {
        super();
    }

    /**
     * Constructs an {@code GoogleException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     */
    public GoogleException(String message) {
        super(message);
    }

    /**
     * Constructs an {@code GoogleException} with the specified detail message
     * and cause.
     *
     * <p> Note that the detail message associated with {@code cause} is
     * <i>not</i> automatically incorporated into this exception's detail
     * message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     *
     * @param cause
     *        The cause (which is saved for later retrieval by the
     *        {@link #getCause()} method).  (A null value is permitted,
     *        and indicates that the cause is nonexistent or unknown.)
     *
     * @since 1.6
     */
    public GoogleException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructs an {@code GoogleException} with the specified detail message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     *
     * @param variables
     *        Array of variables to substitute into the message
     */
    public GoogleException(String message, Object... variables) {
        super(String.format(message, variables));
    }

    /**
     * Constructs an {@code GoogleException} with the specified detail message
     * and cause.
     *
     * <p> Note that the detail message associated with {@code cause} is
     * <i>not</i> automatically incorporated into this exception's detail
     * message.
     *
     * @param message
     *        The detail message (which is saved for later retrieval
     *        by the {@link #getMessage()} method)
     *
     * @param cause
     *        The cause (which is saved for later retrieval by the
     *        {@link #getCause()} method).  (A null value is permitted,
     *        and indicates that the cause is nonexistent or unknown.)
     *
     * @param variables
     *        Array of variables to substitute into the message
     *
     * @since 1.6
     */
    public GoogleException(String message, Throwable cause, Object... variables) {
        super(String.format(message, variables), cause);
    }

    /**
     * Constructs an {@code GoogleException} with the specified cause and a
     * detail message of {@code (cause==null ? null : cause.toString())}
     * (which typically contains the class and detail message of {@code cause}).
     * This constructor is useful for IO exceptions that are little more
     * than wrappers for other throwables.
     *
     * @param cause
     *        The cause (which is saved for later retrieval by the
     *        {@link #getCause()} method).  (A null value is permitted,
     *        and indicates that the cause is nonexistent or unknown.)
     *
     * @since 1.6
     */
    public GoogleException(Throwable cause) {
        super(cause);
    }
}
