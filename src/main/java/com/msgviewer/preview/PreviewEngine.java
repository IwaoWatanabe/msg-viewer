package com.msgviewer.preview;

import com.msgviewer.core.AttachmentData;
import org.mozilla.universalchardet.UniversalDetector;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.FileHeader;
import jakarta.mail.Session;
import jakarta.mail.internet.MimeMessage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * Central engine for rendering file previews.
 * Supports images, text, PDF, Word, Excel, PowerPoint, ZIP, and EML.
 */
public class PreviewEngine {

    private static final Logger log = LoggerFactory.getLogger(PreviewEngine.class);
    private static final long ZIP_MAX_TOTAL_SIZE = 500L * 1024 * 1024; // 500MB
    private static final int ZIP_MAX_FILE_COUNT = 1000;

    /**
     * Generates a preview for the given attachment data.
     */
    public PreviewResult preview(AttachmentData attachment) {
        if (attachment.getData() == null || attachment.getData().length == 0) {
            return PreviewResult.unsupported(attachment.getFileName(), attachment.getMimeType());
        }

        String mime = attachment.getMimeType() != null ? attachment.getMimeType().toLowerCase() : "";
        String name = attachment.getFileName() != null ? attachment.getFileName().toLowerCase() : "";

        try {
            if (mime.startsWith("image/")) {
                return PreviewResult.image(attachment.getData());
            }
            if (mime.startsWith("text/") || name.endsWith(".txt") || name.endsWith(".csv")
                    || name.endsWith(".log") || name.endsWith(".xml") || name.endsWith(".json")) {
                return previewText(attachment.getData());
            }
            if (mime.equals("application/pdf") || name.endsWith(".pdf")) {
                return previewPdf(attachment.getData(), 0);
            }
            if (name.endsWith(".doc") && !name.endsWith(".docx")) {
                return previewDocLegacy(attachment.getData());
            }
            if (name.endsWith(".docx") || mime.contains("wordprocessingml")) {
                return previewDocx(attachment.getData());
            }
            if (name.endsWith(".xls") && !name.endsWith(".xlsx")) {
                return previewXlsLegacy(attachment.getData());
            }
            if (name.endsWith(".xlsx") || mime.contains("spreadsheetml")) {
                return previewXlsx(attachment.getData());
            }
            if (name.endsWith(".ppt") && !name.endsWith(".pptx")) {
                return previewPptLegacy(attachment.getData(), 0);
            }
            if (name.endsWith(".pptx") || mime.contains("presentationml")) {
                return previewPptx(attachment.getData(), 0);
            }
            if (name.endsWith(".eml") || mime.equals("message/rfc822")) {
                return previewEml(attachment.getData());
            }
        } catch (Exception e) {
            log.error("プレビュー生成エラー: {}", attachment.getFileName(), e);
            return PreviewResult.html(errorHtml("プレビューの生成中にエラーが発生しました", e.getMessage()));
        }

        return PreviewResult.unsupported(attachment.getFileName(), attachment.getMimeType());
    }

    /**
     * Generate a specific page of a paged file (PDF/PPT).
     */
    public PreviewResult previewPage(AttachmentData attachment, int page) {
        String name = attachment.getFileName() != null ? attachment.getFileName().toLowerCase() : "";
        try {
            if (name.endsWith(".pdf")) return previewPdf(attachment.getData(), page);
            if (name.endsWith(".ppt") && !name.endsWith(".pptx")) return previewPptLegacy(attachment.getData(), page);
            if (name.endsWith(".pptx")) return previewPptx(attachment.getData(), page);
        } catch (Exception e) {
            log.error("ページプレビューエラー: {}", attachment.getFileName(), e);
        }
        return PreviewResult.unsupported(attachment.getFileName(), attachment.getMimeType());
    }

    // ===================== TEXT =====================

    private PreviewResult previewText(byte[] data) {
        String charset = detectCharset(data);
        String text = new String(data, Charset.forName(charset));
        return PreviewResult.text(text);
    }

    private String detectCharset(byte[] data) {
        UniversalDetector detector = new UniversalDetector(null);
        detector.handleData(data, 0, Math.min(data.length, 8192));
        detector.dataEnd();
        String detected = detector.getDetectedCharset();
        detector.reset();
        return (detected != null) ? detected : StandardCharsets.UTF_8.name();
    }

    // ===================== PDF =====================

    private PreviewResult previewPdf(byte[] data, int page) throws IOException {
        try (PDDocument doc = Loader.loadPDF(data)) {
            int pageCount = doc.getNumberOfPages();
            if (page < 0 || page >= pageCount) page = 0;

            PDFRenderer renderer = new PDFRenderer(doc);
            BufferedImage img = renderer.renderImageWithDPI(page, 150);
            byte[] pngBytes = toImageBytes(img, "png");
            return PreviewResult.pagedImage(pngBytes, pageCount, page);
        }
    }

    // ===================== WORD =====================

    private PreviewResult previewDocLegacy(byte[] data) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             HWPFDocument doc = new HWPFDocument(is)) {

            org.w3c.dom.Document htmlDoc = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder().newDocument();
            WordToHtmlConverter converter = new WordToHtmlConverter(htmlDoc);
            converter.processDocument(doc);

            String html = domToString(converter.getDocument());
            return PreviewResult.html(wrapHtml(html));
        }
    }

    private PreviewResult previewDocx(byte[] data) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             XWPFDocument doc = new XWPFDocument(is)) {

            StringBuilder html = new StringBuilder();
            html.append("<html><head><meta charset='UTF-8'><style>")
                .append("body{font-family:sans-serif;padding:20px;line-height:1.6;}")
                .append("table{border-collapse:collapse;width:100%;}")
                .append("td,th{border:1px solid #ddd;padding:8px;}")
                .append("</style></head><body>");

            doc.getParagraphs().forEach(p -> {
                String style = p.getStyle();
                if (style != null && style.startsWith("Heading")) {
                    int level = Math.min(6, Integer.parseInt(style.replace("Heading", "").trim()));
                    html.append("<h").append(level).append(">")
                        .append(escapeHtml(p.getText()))
                        .append("</h").append(level).append(">");
                } else {
                    html.append("<p>").append(escapeHtml(p.getText())).append("</p>");
                }
            });

            doc.getTables().forEach(table -> {
                html.append("<table>");
                table.getRows().forEach(row -> {
                    html.append("<tr>");
                    row.getTableCells().forEach(cell ->
                        html.append("<td>").append(escapeHtml(cell.getText())).append("</td>")
                    );
                    html.append("</tr>");
                });
                html.append("</table>");
            });

            html.append("</body></html>");
            return PreviewResult.html(html.toString());
        }
    }

    // ===================== EXCEL =====================

    private PreviewResult previewXlsLegacy(byte[] data) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             HSSFWorkbook wb = new HSSFWorkbook(is)) {
            return previewWorkbook(wb);
        }
    }

    private PreviewResult previewXlsx(byte[] data) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             XSSFWorkbook wb = new XSSFWorkbook(is)) {
            return previewWorkbook(wb);
        }
    }

    private PreviewResult previewWorkbook(Workbook wb) {
        StringBuilder html = new StringBuilder();
        html.append("<html><head><meta charset='UTF-8'><style>")
            .append("body{font-family:sans-serif;padding:10px;}")
            .append("h2{color:#333;border-bottom:2px solid #4a90d9;padding-bottom:5px;}")
            .append("table{border-collapse:collapse;width:100%;margin-bottom:20px;}")
            .append("td,th{border:1px solid #ccc;padding:6px 10px;text-align:left;}")
            .append("th{background:#4a90d9;color:white;}")
            .append("tr:nth-child(even){background:#f5f5f5;}")
            .append("</style></head><body>");

        DataFormatter formatter = new DataFormatter();

        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            Sheet sheet = wb.getSheetAt(s);
            html.append("<h2>").append(escapeHtml(sheet.getSheetName())).append("</h2>");
            html.append("<table>");

            int maxRows = Math.min(sheet.getLastRowNum() + 1, 500); // limit rows
            for (int r = 0; r < maxRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                html.append("<tr>");
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String tag = (r == 0) ? "th" : "td";
                    html.append("<").append(tag).append(">")
                        .append(escapeHtml(formatter.formatCellValue(cell)))
                        .append("</").append(tag).append(">");
                }
                html.append("</tr>");
            }

            html.append("</table>");
        }

        html.append("</body></html>");
        return PreviewResult.html(html.toString());
    }

    // ===================== POWERPOINT =====================

    private PreviewResult previewPptLegacy(byte[] data, int page) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             HSLFSlideShow ppt = new HSLFSlideShow(is)) {

            List<HSLFSlide> slides = ppt.getSlides();
            if (slides.isEmpty()) return PreviewResult.text("(スライドがありません)");
            if (page < 0 || page >= slides.size()) page = 0;

            Dimension dim = ppt.getPageSize();
            BufferedImage img = new BufferedImage(dim.width, dim.height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = img.createGraphics();
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, dim.width, dim.height);
            slides.get(page).draw(g);
            g.dispose();

            return PreviewResult.pagedImage(toImageBytes(img, "png"), slides.size(), page);
        }
    }

    private PreviewResult previewPptx(byte[] data, int page) throws Exception {
        try (InputStream is = new ByteArrayInputStream(data);
             XMLSlideShow pptx = new XMLSlideShow(is)) {

            List<XSLFSlide> slides = pptx.getSlides();
            if (slides.isEmpty()) return PreviewResult.text("(スライドがありません)");
            if (page < 0 || page >= slides.size()) page = 0;

            Dimension dim = pptx.getPageSize();
            BufferedImage img = new BufferedImage(dim.width, dim.height, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g = img.createGraphics();
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, dim.width, dim.height);
            slides.get(page).draw(g);
            g.dispose();

            return PreviewResult.pagedImage(toImageBytes(img, "png"), slides.size(), page);
        }
    }

    // ===================== EML =====================

    private PreviewResult previewEml(byte[] data) throws Exception {
        Session session = Session.getDefaultInstance(new Properties());
        try (InputStream is = new ByteArrayInputStream(data)) {
            MimeMessage msg = new MimeMessage(session, is);

            StringBuilder html = new StringBuilder();
            html.append("<html><head><meta charset='UTF-8'><style>")
                .append("body{font-family:sans-serif;padding:20px;}")
                .append(".meta{background:#f0f4f8;padding:15px;border-radius:8px;margin-bottom:15px;}")
                .append(".meta dt{font-weight:bold;color:#555;}")
                .append(".meta dd{margin:0 0 8px 0;}")
                .append("</style></head><body>");

            html.append("<dl class='meta'>");
            html.append("<dt>件名</dt><dd>").append(escapeHtml(str(msg.getSubject()))).append("</dd>");
            html.append("<dt>差出人</dt><dd>").append(escapeHtml(str(msg.getFrom()))).append("</dd>");
            html.append("<dt>宛先</dt><dd>").append(escapeHtml(str(msg.getAllRecipients()))).append("</dd>");
            html.append("<dt>日時</dt><dd>").append(msg.getSentDate() != null ? msg.getSentDate().toString() : "").append("</dd>");
            html.append("</dl>");

            // Try to get body content
            Object content = msg.getContent();
            if (content instanceof String) {
                html.append("<div>").append(escapeHtml((String) content)).append("</div>");
            } else {
                html.append("<div>(本文を抽出できませんでした)</div>");
            }

            html.append("</body></html>");
            return PreviewResult.html(html.toString());
        }
    }

    // ===================== ZIP (metadata only) =====================

    /**
     * Returns a tree-view HTML of the ZIP contents.
     * Actual extraction with password is handled by ZipHandler.
     */
    public PreviewResult previewZipContents(File zipFile, boolean isEncrypted) throws IOException {
        try (ZipFile zf = new ZipFile(zipFile)) {
            List<FileHeader> headers = zf.getFileHeaders();

            StringBuilder html = new StringBuilder();
            html.append("<html><head><meta charset='UTF-8'><style>")
                .append("body{font-family:sans-serif;padding:15px;}")
                .append(".lock{color:#e74c3c;}")
                .append("ul{list-style-type:none;padding-left:18px;}")
                .append("li{padding:3px 0;}")
                .append("li::before{content:'📄 ';}.dir::before{content:'📁 ';}")
                .append("</style></head><body>");

            if (isEncrypted) {
                html.append("<p class='lock'>🔒 パスワードで保護されています</p>");
            }
            html.append("<h3>ZIPファイルの内容</h3><ul>");

            for (FileHeader fh : headers) {
                String cls = fh.isDirectory() ? " class='dir'" : "";
                html.append("<li").append(cls).append(">")
                    .append(escapeHtml(fh.getFileName()))
                    .append("</li>");
            }
            html.append("</ul></body></html>");
            return PreviewResult.html(html.toString());
        }
    }

    // ===================== Utilities =====================

    private byte[] toImageBytes(BufferedImage img, String format) throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            ImageIO.write(img, format, baos);
            return baos.toByteArray();
        }
    }

    private String domToString(org.w3c.dom.Document doc) throws Exception {
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        StringWriter sw = new StringWriter();
        transformer.transform(new DOMSource(doc), new StreamResult(sw));
        return sw.toString();
    }

    private String wrapHtml(String body) {
        return "<html><head><meta charset='UTF-8'><style>body{font-family:sans-serif;padding:20px;line-height:1.6;}</style></head><body>"
                + body + "</body></html>";
    }

    private String errorHtml(String title, String detail) {
        return "<html><head><meta charset='UTF-8'><style>"
                + "body{font-family:sans-serif;padding:20px;color:#333;}"
                + ".error{background:#fee;border:1px solid #fcc;padding:15px;border-radius:8px;}"
                + "</style></head><body><div class='error'>"
                + "<h3>⚠ " + escapeHtml(title) + "</h3>"
                + "<p>" + escapeHtml(detail) + "</p>"
                + "</div></body></html>";
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;").replace("'", "&#39;");
    }

    private String str(Object obj) {
        if (obj == null) return "";
        if (obj instanceof Object[]) return Arrays.toString((Object[]) obj);
        return obj.toString();
    }
}
