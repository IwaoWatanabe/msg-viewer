package com.msgviewer;

import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.FileHeader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.List;

/**
 * Handles ZIP file operations including encrypted ZIP support.
 * Includes ZIP bomb protection via size and file count limits.
 */
class ZipHandler {

    private static final Logger log = LoggerFactory.getLogger(ZipHandler.class);
    private static final long MAX_TOTAL_SIZE = 500L * 1024 * 1024; // 500MB
    private static final int MAX_FILE_COUNT = 1000;

    /**
     * Checks if a ZIP file is encrypted.
     */
    public boolean isEncrypted(byte[] zipData) throws IOException {
        Path tempFile = writeTempZip(zipData);
        try (ZipFile zf = new ZipFile(tempFile.toFile())) {
            return zf.isEncrypted();
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    /**
     * Lists the file headers (entries) in a ZIP file.
     */
    public List<FileHeader> listEntries(byte[] zipData) throws IOException {
        Path tempFile = writeTempZip(zipData);
        try (ZipFile zf = new ZipFile(tempFile.toFile())) {
            return zf.getFileHeaders();
        } finally {
            Files.deleteIfExists(tempFile);
        }
    }

    /**
     * Extracts a ZIP file to a temporary directory with ZIP bomb protection.
     *
     * @param zipData   raw ZIP bytes
     * @param password  password for encrypted ZIPs (null for non-encrypted)
     * @return path to the temporary extraction directory
     */
    public Path extract(byte[] zipData, char[] password) throws IOException {
        Path tempZip = writeTempZip(zipData);
        Path extractDir = Files.createTempDirectory("msg-viewer-zip-");

        try {
            ZipFile zf = new ZipFile(tempZip.toFile());
            if (password != null) {
                zf.setPassword(password);
            }

            // Validate before extraction (ZIP bomb check)
            long totalUncompressed = 0;
            int fileCount = 0;
            for (FileHeader fh : zf.getFileHeaders()) {
                if (!fh.isDirectory()) {
                    totalUncompressed += fh.getUncompressedSize();
                    fileCount++;
                }
                if (totalUncompressed > MAX_TOTAL_SIZE) {
                    throw new IOException("ZIPボム検知: 展開後のサイズが上限 (500MB) を超えます");
                }
                if (fileCount > MAX_FILE_COUNT) {
                    throw new IOException("ZIPボム検知: ファイル数が上限 (1000) を超えます");
                }
            }

            zf.extractAll(extractDir.toString());
            return extractDir;

        } catch (Exception e) {
            // Cleanup on failure
            deleteDirectory(extractDir);
            throw new IOException("ZIP展開に失敗しました: " + e.getMessage(), e);
        } finally {
            Files.deleteIfExists(tempZip);
        }
    }

    /**
     * Extracts a single entry from a ZIP file.
     */
    public byte[] extractEntry(byte[] zipData, String entryName, char[] password) throws IOException {
        Path tempZip = writeTempZip(zipData);
        Path extractDir = Files.createTempDirectory("msg-viewer-zipentry-");

        try {
            ZipFile zf = new ZipFile(tempZip.toFile());
            if (password != null) {
                zf.setPassword(password);
            }
            zf.extractFile(entryName, extractDir.toString());

            Path extracted = extractDir.resolve(entryName);
            if (Files.exists(extracted)) {
                return Files.readAllBytes(extracted);
            }
            throw new IOException("エントリが見つかりません: " + entryName);
        } finally {
            deleteDirectory(extractDir);
            Files.deleteIfExists(tempZip);
        }
    }

    private Path writeTempZip(byte[] data) throws IOException {
        Path tempFile = Files.createTempFile("msg-viewer-", ".zip");
        Files.write(tempFile, data);
        return tempFile;
    }

    /**
     * Recursively deletes a directory.
     */
    public static void deleteDirectory(Path dir) {
        if (dir == null || !Files.exists(dir)) return;
        try {
            Files.walkFileTree(dir, new SimpleFileVisitor<>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Files.deleteIfExists(file);
                    return FileVisitResult.CONTINUE;
                }
                @Override
                public FileVisitResult postVisitDirectory(Path d, IOException exc) throws IOException {
                    Files.deleteIfExists(d);
                    return FileVisitResult.CONTINUE;
                }
            });
        } catch (IOException e) {
            log.warn("ディレクトリの削除に失敗: {}", dir, e);
        }
    }
}
