/*
 *
 * Copyright (c) 2025, Pivotal Solutions Ltd and/or its affiliates. All rights reserved.
 * Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 */
package com.pivotal.google.docs;

import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.pivotal.google.GoogleServiceFactory;
import com.pivotal.utils.EzGdocs4jException;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.io.FilenameFilter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

/**
 * A wrapper around a Google File
 */
@SuppressWarnings({"unused", "BooleanMethodIsAlwaysInverted"})
@Slf4j
public class GoogleFile {

    @Getter
    private final String fileId;

    @Getter(AccessLevel.PROTECTED)
    private com.google.api.services.drive.model.File file = null;

    /**
     * Creates an instance of the root folder
     *
     * @throws com.pivotal.google.docs.GoogleException If the file doesn't exist or no permission to access it
     */
    public GoogleFile() throws com.pivotal.google.docs.GoogleException {
        this.fileId = "root";
        initFile();
    }

    /**
     * Creates an instance of the file using the specified ID
     *
     * @param fileId Unique ID of the file
     * @throws com.pivotal.google.docs.GoogleException If the file doesn't exist or no permission to access it
     */
    public GoogleFile(String fileId) throws com.pivotal.google.docs.GoogleException {
        this.fileId = fileId;
        initFile();
    }

    /**
     * Create a folder in the root folder
     *
     * @param name Name to give the folder
     * @return GoogleFile of type folder
     * @throws com.pivotal.google.docs.GoogleException If it cannot create the folder
     */
    public static GoogleFile createFolder(String name) throws com.pivotal.google.docs.GoogleException {
        return createFolder(name, null);
    }

    /**
     * Create a folder in the given folder or root if null
     *
     * @param name     Name to give the folder
     * @param folderId Id of the destination folder
     * @return GoogleFile of type folder
     * @throws com.pivotal.google.docs.GoogleException If it cannot create the folder
     */
    public static GoogleFile createFolder(String name, String folderId) throws com.pivotal.google.docs.GoogleException {
        File fileMetadata = new File();
        fileMetadata.setName(name);
        fileMetadata.setMimeType("application/vnd.google-apps.folder");
        if (folderId != null) {
            fileMetadata.setParents(Collections.singletonList(folderId));
        }
        try {
            File file = GoogleServiceFactory.getFilesService().create(fileMetadata)
                    .setFields("id")
                    .execute();
            return new GoogleFile(file.getId());
        }
        catch (IOException e) {
            throw new com.pivotal.google.docs.GoogleException("Failed to create folder [%s] - %s", name, e.getMessage());
        }
    }

    /**
     * Finds the first file in the folder that has the specified name
     *
     * @param fileName Name of the file to get
     * @param folderId ID of the folder to search
     * @return Null if the file doesn't exist
     * @throws com.pivotal.google.docs.GoogleException If the folder doesn't exist
     */
    public static GoogleFile getFileFromFolderByName(String fileName, String folderId) throws com.pivotal.google.docs.GoogleException {
        GoogleFile folder = new GoogleFile(folderId);
        List<GoogleFile> files = folder.list((dir, name) -> name.equalsIgnoreCase(fileName));
        if (!files.isEmpty()) {
            return files.get(0);
        }
        return null;
    }

    /**
     * Returns the name/title of the file
     *
     * @return File name
     */
    public String getName() {
        return file.getName();
    }

    /**
     * Returns the list of parents that reference this file
     *
     * @return List of parents
     */
    public List<GoogleFile> getParents() {
        List<GoogleFile> ret = new ArrayList<>();
        for (String parentId : file.getParents()) {
            try {
                ret.add(new GoogleFile(parentId));
            }
            catch (com.pivotal.google.docs.GoogleException e) {
                log.warn("Cannot get parent file {} - {}", parentId, e.getMessage());
                throw new EzGdocs4jException(e);
            }
        }
        return ret;
    }

    /**
     * Returns the URL that can be used to access this file in Drive
     *
     * @return URL of this file
     * @throws MalformedURLException If the URL cannot be created
     */
    public URL toURL() throws MalformedURLException {
        return new URL(file.getWebViewLink());
    }

    /**
     * Returns true if this file/folder is writable
     *
     * @return True if writable
     */
    public boolean canWrite() {
        return file.getCapabilities().getCanModifyContent();
    }

    /**
     * Returns true if this is a folder
     *
     * @return True if this is a folder
     */
    public boolean isFolder() {
        return file.getMimeType().equals("application/vnd.google-apps.folder");
    }

    /**
     * Returns true iof this is a standard file
     *
     * @return True if this is a file, not a folder
     */
    public boolean isFile() {
        return !isFolder();
    }

    public Date lastModified() {
        return new Date(file.getModifiedTime().getValue());
    }

    /**
     * Returns the Date whe the file was created
     *
     * @return Creation date
     */
    public Date created() {
        return new Date(file.getCreatedTime().getValue());
    }

    /**
     * Returns the size in bytes of the file
     *
     * @return Number of bytes
     */
    public long size() {
        return file.getSize();
    }


    /**
     * Lists all the files under this file.
     * If this file isn't a folder, it returns an empty list
     *
     * @return List of GoogleFiles
     */
    public List<GoogleFile> list() {
        return list(null);
    }

    /**
     * Lists all the files under this file using the specified filter.
     * If this file isn't a folder, it returns an empty list
     *
     * @param filter Filename filter to use
     * @return List of GoogleFiles
     */
    public List<GoogleFile> list(FilenameFilter filter) {
        List<GoogleFile> ret = new ArrayList<>();
        try {
            FileList result;
            do {
                // Get a page of results
                result = GoogleServiceFactory.getFilesService().list()
                        .setPageSize(100)
                        .setQ(String.format("'%s' in parents", fileId))
                        .setOrderBy("name")
                        .setFields("nextPageToken, files(id, name)")
                        .execute();

                // Add them to the list
                for (File foundFile : result.getFiles()) {
                    try {
                        if (filter == null || filter.accept(null, foundFile.getName())) {
                            ret.add(new GoogleFile(foundFile.getId()));
                        }
                    }
                    catch (com.pivotal.google.docs.GoogleException e) {
                        log.error(e.getMessage());
                    }
                }
            } while (result.getNextPageToken() != null);
        }
        catch (IOException e) {
            log.debug("Cannot list files from {} - {}", fileId, e.getMessage());
        }
        return ret;
    }

    /**
     * Renames the file to the new name
     *
     * @param name New name to give the file
     * @throws com.pivotal.google.docs.GoogleException if there is a problem renaming and then accessing file
     */
    public void renameTo(String name) throws com.pivotal.google.docs.GoogleException {
        try {
            File localFile = new File();
            localFile.setName(name);
            GoogleServiceFactory.getFilesService().update(fileId, localFile)
                    .setFields("name")
                    .execute();
            initFile();
        }
        catch (IOException e) {
            throw new com.pivotal.google.docs.GoogleException("Failed to rename file [%s] (%s) to [%s] - %s", file.getName(), fileId, name, e.getMessage());
        }
    }

    /**
     * Moves the file from whatever parent it has now, to a new parent folder
     *
     * @param folderId ID of the folder to move to
     * @throws com.pivotal.google.docs.GoogleException If it cannot be moved
     */
    public void move(String folderId) throws com.pivotal.google.docs.GoogleException {

        // Check for rubbish
        if (folderId != null && !folderId.isEmpty()) {
            try {

                // Is this actually a folder
                GoogleFile folder = new GoogleFile(folderId);
                if (!folder.isFolder()) {
                    throw new com.pivotal.google.docs.GoogleException("Destination %s [%s] isn't a folder", folder.getName(), folderId);
                }

                // Retrieve the existing parents to remove
                GoogleFile googleFile = new GoogleFile(fileId);
                String previousParents = String.join(",", googleFile.getFile().getParents());
                try {

                    // Move the file to the folder
                    GoogleServiceFactory.getFilesService().update(fileId, null)
                            .setAddParents(folderId)
                            .setRemoveParents(previousParents)
                            .setFields("id, parents")
                            .execute();
                }
                catch (IOException e) {
                    throw new com.pivotal.google.docs.GoogleException("Cannot move file %s to folder %s", e, fileId, folderId);
                }

            }
            catch (com.pivotal.google.docs.GoogleException e) {
                throw new com.pivotal.google.docs.GoogleException("Cannot open destination folder %s", folderId);
            }
        }
    }

    /**
     * Attempts to delete this file
     *
     * @throws com.pivotal.google.docs.GoogleException If it cannot be deleted
     */
    public void delete() throws com.pivotal.google.docs.GoogleException {
        try {
            GoogleServiceFactory.getFilesService().delete(fileId).execute();
        }
        catch (IOException e) {
            throw new com.pivotal.google.docs.GoogleException("Failed to delete file [%s] (%s) - %s", file.getName(), fileId, e.getMessage());
        }
    }


    /**
     * Initialises the file as part of the object instantiation
     *
     * @throws com.pivotal.google.docs.GoogleException If the file cannot be accessed
     */
    private void initFile() throws com.pivotal.google.docs.GoogleException {
        try {
            file = GoogleServiceFactory.getFilesService().get(fileId)
                    .setFields("kind,id,name,mimeType,starred,trashed,spaces,webViewLink,createdTime,modifiedTime,owners,lastModifyingUser,capabilities,size,quotaBytesUsed,permissions")
                    .execute();
        }
        catch (IOException e) {
            throw new com.pivotal.google.docs.GoogleException("Cannot get file %s", e, fileId);
        }
    }

    /**
     * Convenience method to check if a file exists without raising an exception
     *
     * @param fileId File ID to check
     * @return True if it can be accessed/opened
     */
    public static boolean exists(String fileId) {
        try {
            GoogleServiceFactory.getFilesService().get(fileId)
                    .setFields("id")
                    .execute();
            return true;
        }
        catch (IOException e) {
            log.debug("Cannot get file {}", fileId, e);
            return false;
        }
    }

}
