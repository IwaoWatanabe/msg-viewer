package com.msgviewer.core;

/**
 * Represents a single attachment extracted from an MSG file.
 */
public class AttachmentData {
    private String fileName;
    private String mimeType;
    private byte[] data;
    private long size;

    public AttachmentData() {}

    public AttachmentData(String fileName, String mimeType, byte[] data) {
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.data = data;
        this.size = (data != null) ? data.length : 0;
    }

    public String getFileName() { return fileName; }
    public void setFileName(String fileName) { this.fileName = fileName; }

    public String getMimeType() { return mimeType; }
    public void setMimeType(String mimeType) { this.mimeType = mimeType; }

    public byte[] getData() { return data; }
    public void setData(byte[] data) {
        this.data = data;
        this.size = (data != null) ? data.length : 0;
    }

    public long getSize() { return size; }

    /**
     * Returns a human-readable file size string.
     */
    public String getFormattedSize() {
        if (size < 1024) return size + " B";
        if (size < 1024 * 1024) return String.format("%.1f KB", size / 1024.0);
        return String.format("%.1f MB", size / (1024.0 * 1024.0));
    }
}
