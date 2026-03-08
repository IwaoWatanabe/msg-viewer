package com.msgviewer.preview;

/**
 * Result of rendering a file for preview in the SWT GUI.
 * Can contain HTML content (for Browser widget), image bytes, or plain text.
 */
public class PreviewResult {

    public enum Type {
        HTML,       // Render in SWT Browser widget
        IMAGE,      // Render as SWT Image
        TEXT,       // Render in StyledText widget
        UNSUPPORTED // Show file info + save button
    }

    private final Type type;
    private final String content;      // HTML or text content
    private final byte[] imageData;    // Image bytes (for IMAGE type)
    private final String fileName;
    private final String mimeType;
    private final int pageCount;       // For PDF/PPT paging
    private final int currentPage;

    private PreviewResult(Type type, String content, byte[] imageData,
                         String fileName, String mimeType, int pageCount, int currentPage) {
        this.type = type;
        this.content = content;
        this.imageData = imageData;
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.pageCount = pageCount;
        this.currentPage = currentPage;
    }

    public static PreviewResult html(String htmlContent) {
        return new PreviewResult(Type.HTML, htmlContent, null, null, null, 0, 0);
    }

    public static PreviewResult image(byte[] data) {
        return new PreviewResult(Type.IMAGE, null, data, null, null, 0, 0);
    }

    public static PreviewResult text(String textContent) {
        return new PreviewResult(Type.TEXT, textContent, null, null, null, 0, 0);
    }

    public static PreviewResult pagedImage(byte[] data, int pageCount, int currentPage) {
        return new PreviewResult(Type.IMAGE, null, data, null, null, pageCount, currentPage);
    }

    public static PreviewResult unsupported(String fileName, String mimeType) {
        return new PreviewResult(Type.UNSUPPORTED, null, null, fileName, mimeType, 0, 0);
    }

    public Type getType() { return type; }
    public String getContent() { return content; }
    public byte[] getImageData() { return imageData; }
    public String getFileName() { return fileName; }
    public String getMimeType() { return mimeType; }
    public int getPageCount() { return pageCount; }
    public int getCurrentPage() { return currentPage; }
}
