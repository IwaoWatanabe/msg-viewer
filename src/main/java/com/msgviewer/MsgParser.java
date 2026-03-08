package com.msgviewer;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;
import org.apache.tika.Tika;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * Parses Outlook MSG files using Apache POI HSMF.
 * Extracts metadata, body text/HTML, and attachments.
 */
class MsgParser {

    private static final Logger log = LoggerFactory.getLogger(MsgParser.class);
    private final Tika tika = new Tika();

    /**
     * Parses the given MSG file and returns structured data.
     *
     * @param msgFile the .msg file to parse
     * @return parsed MsgData
     * @throws IOException if the file cannot be read or parsed
     */
    public MsgData parse(File msgFile) throws IOException {
        MsgData data = new MsgData();

        try (MAPIMessage msg = new MAPIMessage(msgFile)) {
            // Metadata
            data.setSubject(safeGetString(() -> msg.getSubject()));
            data.setFrom(safeGetString(() -> msg.getDisplayFrom()));
            data.setTo(safeGetString(() -> msg.getDisplayTo()));
            data.setCc(safeGetString(() -> msg.getDisplayCC()));
            data.setDate(safeGetDate(msg));

            // Body
            data.setBodyText(safeGetString(() -> msg.getTextBody()));
            data.setBodyHtml(safeGetString(() -> msg.getHtmlBody()));

            // Attachments
            data.setAttachments(extractAttachments(msg));

        } catch (Exception e) {
            throw new IOException("MSGファイルの解析に失敗しました: " + msgFile.getName(), e);
        }

        return data;
    }

    private List<AttachmentData> extractAttachments(MAPIMessage msg) {
        List<AttachmentData> attachments = new ArrayList<>();
        AttachmentChunks[] chunks = msg.getAttachmentFiles();

        if (chunks == null) return attachments;

        for (AttachmentChunks chunk : chunks) {
            try {
                String fileName = null;
                if (chunk.getAttachLongFileName() != null) {
                    fileName = chunk.getAttachLongFileName().getValue();
                } else if (chunk.getAttachFileName() != null) {
                    fileName = chunk.getAttachFileName().getValue();
                }

                byte[] fileData = null;
                if (chunk.getAttachData() != null) {
                    fileData = chunk.getAttachData().getValue();
                }

                if (fileName == null || fileName.isBlank()) {
                    fileName = "attachment_" + attachments.size();
                }

                // Detect MIME type
                String mimeType = "application/octet-stream";
                if (fileData != null && fileData.length > 0) {
                    try {
                        mimeType = tika.detect(new ByteArrayInputStream(fileData), fileName);
                    } catch (Exception e) {
                        log.warn("MIME type detection failed for: {}", fileName, e);
                    }
                }

                attachments.add(new AttachmentData(fileName, mimeType, fileData));
            } catch (Exception e) {
                log.warn("添付ファイルの抽出に失敗: {}", e.getMessage());
            }
        }
        return attachments;
    }

    private String safeGetDate(MAPIMessage msg) {
        try {
            if (msg.getMessageDate() != null) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                return sdf.format(msg.getMessageDate().getTime());
            }
        } catch (ChunkNotFoundException e) {
            // ignore
        }
        return "";
    }

    @FunctionalInterface
    private interface StringSupplier {
        String get() throws ChunkNotFoundException;
    }

    private String safeGetString(StringSupplier supplier) {
        try {
            String value = supplier.get();
            return (value != null) ? value : "";
        } catch (ChunkNotFoundException e) {
            return "";
        }
    }
}
