package com.msgviewer;

import java.time.LocalDateTime;

/**
 * Represents a single comment attached to a MSG file or its attachment.
 */
class Comment {
    private long id;
    private String fileHash;       // Hash key identifying the MSG or attachment
    private String fileName;       // Display name of the target file
    private String text;           // Comment content
    private LocalDateTime createdAt;
    private LocalDateTime updatedAt;

    public Comment() {}

    public Comment(String fileHash, String fileName, String text) {
        this.fileHash = fileHash;
        this.fileName = fileName;
        this.text = text;
        this.createdAt = LocalDateTime.now();
        this.updatedAt = this.createdAt;
    }

    public long getId() { return id; }
    public void setId(long id) { this.id = id; }

    public String getFileHash() { return fileHash; }
    public void setFileHash(String fileHash) { this.fileHash = fileHash; }

    public String getFileName() { return fileName; }
    public void setFileName(String fileName) { this.fileName = fileName; }

    public String getText() { return text; }
    public void setText(String text) { this.text = text; }

    public LocalDateTime getCreatedAt() { return createdAt; }
    public void setCreatedAt(LocalDateTime createdAt) { this.createdAt = createdAt; }

    public LocalDateTime getUpdatedAt() { return updatedAt; }
    public void setUpdatedAt(LocalDateTime updatedAt) { this.updatedAt = updatedAt; }
}
