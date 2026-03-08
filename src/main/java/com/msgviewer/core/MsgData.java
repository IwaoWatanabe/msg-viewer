package com.msgviewer.core;

import java.util.List;

/**
 * Represents a parsed Outlook MSG file with metadata, body, and attachments.
 */
public class MsgData {
    private String subject;
    private String from;
    private String to;
    private String cc;
    private String date;
    private String bodyText;
    private String bodyHtml;
    private List<AttachmentData> attachments;

    public MsgData() {}

    public String getSubject() { return subject; }
    public void setSubject(String subject) { this.subject = subject; }

    public String getFrom() { return from; }
    public void setFrom(String from) { this.from = from; }

    public String getTo() { return to; }
    public void setTo(String to) { this.to = to; }

    public String getCc() { return cc; }
    public void setCc(String cc) { this.cc = cc; }

    public String getDate() { return date; }
    public void setDate(String date) { this.date = date; }

    public String getBodyText() { return bodyText; }
    public void setBodyText(String bodyText) { this.bodyText = bodyText; }

    public String getBodyHtml() { return bodyHtml; }
    public void setBodyHtml(String bodyHtml) { this.bodyHtml = bodyHtml; }

    public List<AttachmentData> getAttachments() { return attachments; }
    public void setAttachments(List<AttachmentData> attachments) { this.attachments = attachments; }
}
