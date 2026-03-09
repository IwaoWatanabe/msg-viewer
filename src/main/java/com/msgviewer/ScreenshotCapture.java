package com.msgviewer;

import org.eclipse.swt.SWT;
import org.eclipse.swt.custom.SashForm;
import org.eclipse.swt.custom.StyledText;
import org.eclipse.swt.graphics.*;
import org.eclipse.swt.layout.*;
import org.eclipse.swt.widgets.*;

import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Launches the SWT GUI with test MSG files, captures screenshots, then exits.
 */
public class ScreenshotCapture {

    private static final Path OUT_DIR = Path.of("screenshots");
    private static final Path SAMPLE_DIR = Path.of("sample-msg");

    public static void main(String[] args) throws Exception {
        Files.createDirectories(OUT_DIR);

        Display display = new Display();

        // Capture 1: Initial empty window
        captureEmptyWindow(display);

        // Capture 2: Window with simple MSG loaded
        captureWithMsg(display, "simple_test_msg.msg");

        // Capture 3: Window with attachment MSG loaded
        captureWithMsg(display, "attachment_test_msg.msg");

        display.dispose();
        System.out.println("✅ スクリーンショット保存先: " + OUT_DIR.toAbsolutePath());
    }

    private static void captureEmptyWindow(Display display) throws Exception {
        Shell shell = createMainShell(display);
        shell.open();

        // Process events to render
        processEvents(display, 500);

        // Capture
        saveShellScreenshot(shell, OUT_DIR.resolve("01_initial_window.png").toString());
        System.out.println("📸 01_initial_window.png captured");

        shell.dispose();
    }

    private static void captureWithMsg(Display display, String msgFile) throws Exception {
        Shell shell = createMainShell(display);

        // Parse MSG
        MsgParser parser = new MsgParser();
        MsgData data = parser.parse(SAMPLE_DIR.resolve(msgFile).toFile());

        // Build layout with data
        populateShell(shell, data, display);

        shell.open();
        processEvents(display, 500);

        String name = msgFile.replace(".msg", "");
        saveShellScreenshot(shell, OUT_DIR.resolve("02_" + name + ".png").toString());
        System.out.println("📸 02_" + name + ".png captured");

        shell.dispose();
    }

    private static Shell createMainShell(Display display) {
        Shell shell = new Shell(display);
        shell.setText("MSG Viewer - テスト");
        shell.setSize(900, 650);
        shell.setLayout(new FillLayout());
        return shell;
    }

    private static void populateShell(Shell shell, MsgData data, Display display) {
        SashForm mainSash = new SashForm(shell, SWT.HORIZONTAL);

        // Left: Mail details
        Composite leftPane = new Composite(mainSash, SWT.BORDER);
        leftPane.setLayout(new GridLayout(1, false));

        // Header
        Label headerLabel = new Label(leftPane, SWT.NONE);
        headerLabel.setText("📧 メール詳細");
        Font boldFont = new Font(display, new FontData("", 14, SWT.BOLD));
        headerLabel.setFont(boldFont);
        headerLabel.setLayoutData(new GridData(SWT.FILL, SWT.TOP, true, false));

        // Metadata
        StyledText metaText = new StyledText(leftPane, SWT.READ_ONLY | SWT.WRAP | SWT.V_SCROLL);
        metaText.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, false));
        StringBuilder meta = new StringBuilder();
        meta.append("件名: ").append(data.getSubject()).append("\n");
        meta.append("差出人: ").append(data.getFrom()).append("\n");
        meta.append("宛先: ").append(data.getTo()).append("\n");
        if (data.getCc() != null && !data.getCc().isBlank()) {
            meta.append("CC: ").append(data.getCc()).append("\n");
        }
        meta.append("日時: ").append(data.getDate()).append("\n");
        metaText.setText(meta.toString());

        // Body
        Label bodyLabel = new Label(leftPane, SWT.NONE);
        bodyLabel.setText("本文:");
        StyledText bodyText = new StyledText(leftPane, SWT.READ_ONLY | SWT.WRAP | SWT.V_SCROLL);
        bodyText.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true));
        bodyText.setText(data.getBodyText() != null ? data.getBodyText() : "(本文なし)");

        // Right pane
        Composite rightPane = new Composite(mainSash, SWT.BORDER);
        rightPane.setLayout(new GridLayout(1, false));

        if (!data.getAttachments().isEmpty()) {
            Label attLabel = new Label(rightPane, SWT.NONE);
            attLabel.setText("📎 添付ファイル (" + data.getAttachments().size() + "件)");
            attLabel.setFont(boldFont);

            org.eclipse.swt.widgets.List attList = new org.eclipse.swt.widgets.List(rightPane, SWT.BORDER | SWT.V_SCROLL);
            attList.setLayoutData(new GridData(SWT.FILL, SWT.TOP, true, false));
            for (AttachmentData att : data.getAttachments()) {
                attList.add(att.getFileName() + " (" + att.getFormattedSize() + ")");
            }
            if (attList.getItemCount() > 0) {
                attList.select(0);
            }

            // Preview area
            Label prevLabel = new Label(rightPane, SWT.NONE);
            prevLabel.setText("プレビュー:");
            StyledText previewText = new StyledText(rightPane, SWT.READ_ONLY | SWT.WRAP | SWT.V_SCROLL);
            previewText.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true));

            AttachmentData firstAtt = data.getAttachments().get(0);
            PreviewEngine engine = new PreviewEngine();
            PreviewResult result = engine.preview(firstAtt);
            switch (result.getType()) {
                case TEXT:
                    previewText.setText(result.getContent());
                    break;
                case HTML:
                    previewText.setText("[HTML プレビュー] " + firstAtt.getFileName());
                    break;
                case IMAGE:
                    previewText.setText("[画像プレビュー] " + firstAtt.getFileName());
                    break;
                default:
                    previewText.setText("[非対応形式] " + firstAtt.getFileName());
            }
        } else {
            Label noAtt = new Label(rightPane, SWT.NONE);
            noAtt.setText("添付ファイルなし");
            noAtt.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true));
        }

        // Comment section at bottom of right
        Label commentLabel = new Label(rightPane, SWT.NONE);
        commentLabel.setText("💬 コメント:");
        Text commentInput = new Text(rightPane, SWT.BORDER);
        commentInput.setLayoutData(new GridData(SWT.FILL, SWT.BOTTOM, true, false));
        commentInput.setMessage("コメントを入力...");

        mainSash.setWeights(new int[]{50, 50});
    }

    private static void processEvents(Display display, int durationMs) {
        long end = System.currentTimeMillis() + durationMs;
        while (System.currentTimeMillis() < end) {
            if (!display.readAndDispatch()) {
                display.sleep();
            }
        }
    }

    private static void saveShellScreenshot(Shell shell, String path) {
        Point size = shell.getSize();
        GC gc = new GC(shell);
        Image image = new Image(shell.getDisplay(), size.x, size.y);
        gc.copyArea(image, 0, 0);
        gc.dispose();

        ImageLoader loader = new ImageLoader();
        loader.data = new ImageData[]{image.getImageData()};
        loader.save(path, SWT.IMAGE_PNG);
        image.dispose();
    }
}
