package com.msgviewer;

import org.eclipse.swt.SWT;
import org.eclipse.swt.browser.Browser;
import org.eclipse.swt.custom.SashForm;
import org.eclipse.swt.custom.StackLayout;
import org.eclipse.swt.dnd.*;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.*;
import org.eclipse.swt.layout.*;
import org.eclipse.swt.widgets.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Main application window built with SWT.
 * Layout:
 *   ┌──────────────┬──────────────────┐
 *   │ Mail Detail   │ Attachment Tree  │
 *   │ (Browser)     │                  │
 *   │               ├──────────────────┤
 *   │               │ Preview Area     │
 *   │               │ (Browser/Canvas) │
 *   ├───────────────┴──────────────────┤
 *   │ Comment Panel                     │
 *   └──────────────────────────────────┘
 */
class MainWindow {

    private static final Logger log = LoggerFactory.getLogger(MainWindow.class);
    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    private final Shell shell;
    private final Display display;
    private final MsgParser parser = new MsgParser();
    private final PreviewEngine previewEngine = new PreviewEngine();
    private final ZipHandler zipHandler = new ZipHandler();
    private final CommentRepository commentRepo = new CommentRepository();

    // UI Components
    private Browser mailDetailBrowser;
    private Tree attachmentTree;
    private Composite previewContainer;
    private Browser previewBrowser;
    private Canvas previewCanvas;
    private org.eclipse.swt.widgets.List commentList;
    private Text commentInput;
    private Label statusBar;

    // State
    private MsgData currentMsg;
    private AttachmentData selectedAttachment;
    private String currentMsgHash;
    private String currentAttachmentHash;
    private int currentPage = 0;
    private int totalPages = 0;
    private byte[] currentMsgFileBytes;

    public MainWindow(Shell shell) {
        this.shell = shell;
        this.display = shell.getDisplay();
        createUI();
        setupDragAndDrop();
    }

    private void createUI() {
        shell.setLayout(new FormLayout());

        // Menu bar
        Menu menuBar = new Menu(shell, SWT.BAR);
        shell.setMenuBar(menuBar);
        MenuItem fileMenu = new MenuItem(menuBar, SWT.CASCADE);
        fileMenu.setText("ファイル(&F)");
        Menu fileDropdown = new Menu(shell, SWT.DROP_DOWN);
        fileMenu.setMenu(fileDropdown);

        MenuItem openItem = new MenuItem(fileDropdown, SWT.PUSH);
        openItem.setText("開く(&O)\tCtrl+O");
        openItem.setAccelerator(SWT.MOD1 | 'O');
        openItem.addSelectionListener(new SelectionAdapter() {
            @Override
            public void widgetSelected(SelectionEvent e) { openFileDialog(); }
        });

        new MenuItem(fileDropdown, SWT.SEPARATOR);
        MenuItem exitItem = new MenuItem(fileDropdown, SWT.PUSH);
        exitItem.setText("終了(&X)");
        exitItem.addSelectionListener(new SelectionAdapter() {
            @Override
            public void widgetSelected(SelectionEvent e) { shell.close(); }
        });

        // Main vertical sash: top (content) + bottom (comments)
        SashForm mainSash = new SashForm(shell, SWT.VERTICAL);
        FormData mainSashData = new FormData();
        mainSashData.top = new FormAttachment(0, 0);
        mainSashData.left = new FormAttachment(0, 0);
        mainSashData.right = new FormAttachment(100, 0);
        mainSashData.bottom = new FormAttachment(100, -25);
        mainSash.setLayoutData(mainSashData);

        // Top horizontal sash: left (mail detail) + right (tree + preview)
        SashForm topSash = new SashForm(mainSash, SWT.HORIZONTAL);

        // Left: Mail detail
        mailDetailBrowser = new Browser(topSash, SWT.NONE);
        mailDetailBrowser.setText(welcomeHtml());

        // Right vertical sash: tree + preview
        SashForm rightSash = new SashForm(topSash, SWT.VERTICAL);

        // Attachment tree
        Composite treeContainer = new Composite(rightSash, SWT.NONE);
        treeContainer.setLayout(new GridLayout(1, false));
        Label treeLabel = new Label(treeContainer, SWT.NONE);
        treeLabel.setText("📎 添付ファイル");
        treeLabel.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false));
        attachmentTree = new Tree(treeContainer, SWT.BORDER | SWT.SINGLE);
        attachmentTree.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true));
        attachmentTree.addSelectionListener(new SelectionAdapter() {
            @Override
            public void widgetSelected(SelectionEvent e) { onAttachmentSelected(); }
        });

        // Preview area
        previewContainer = new Composite(rightSash, SWT.NONE);
        previewContainer.setLayout(new StackLayout());
        previewBrowser = new Browser(previewContainer, SWT.NONE);
        previewBrowser.setText("<html><body style='font-family:sans-serif;padding:20px;color:#888;'>"
                + "<p>添付ファイルを選択するとプレビューが表示されます</p></body></html>");

        previewCanvas = new Canvas(previewContainer, SWT.V_SCROLL | SWT.H_SCROLL | SWT.BORDER);
        ((StackLayout) previewContainer.getLayout()).topControl = previewBrowser;

        topSash.setWeights(new int[]{40, 60});
        rightSash.setWeights(new int[]{35, 65});

        // Bottom: Comment panel
        Composite commentPanel = new Composite(mainSash, SWT.NONE);
        commentPanel.setLayout(new GridLayout(1, false));

        Label commentLabel = new Label(commentPanel, SWT.NONE);
        commentLabel.setText("💬 コメント");
        commentLabel.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false));

        commentList = new org.eclipse.swt.widgets.List(commentPanel, SWT.BORDER | SWT.V_SCROLL);
        commentList.setLayoutData(new GridData(SWT.FILL, SWT.FILL, true, true));

        Composite commentInputRow = new Composite(commentPanel, SWT.NONE);
        commentInputRow.setLayout(new GridLayout(3, false));
        commentInputRow.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false));

        commentInput = new Text(commentInputRow, SWT.BORDER);
        commentInput.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false));
        commentInput.setMessage("コメントを入力...");

        Button addBtn = new Button(commentInputRow, SWT.PUSH);
        addBtn.setText("追加");
        addBtn.addSelectionListener(new SelectionAdapter() {
            @Override
            public void widgetSelected(SelectionEvent e) { addComment(); }
        });

        Button deleteBtn = new Button(commentInputRow, SWT.PUSH);
        deleteBtn.setText("削除");
        deleteBtn.addSelectionListener(new SelectionAdapter() {
            @Override
            public void widgetSelected(SelectionEvent e) { deleteSelectedComment(); }
        });

        // Enter key to add comment
        commentInput.addListener(SWT.DefaultSelection, event -> addComment());

        mainSash.setWeights(new int[]{75, 25});

        // Status bar
        statusBar = new Label(shell, SWT.BORDER);
        FormData statusData = new FormData();
        statusData.left = new FormAttachment(0, 0);
        statusData.right = new FormAttachment(100, 0);
        statusData.bottom = new FormAttachment(100, 0);
        statusData.height = 20;
        statusBar.setLayoutData(statusData);
        statusBar.setText("準備完了 — MSGファイルをドラッグ&ドロップするか、ファイル→開く で選択してください");
    }

    private void setupDragAndDrop() {
        DropTarget dropTarget = new DropTarget(shell, DND.DROP_COPY | DND.DROP_DEFAULT);
        dropTarget.setTransfer(new Transfer[]{FileTransfer.getInstance()});
        dropTarget.addDropListener(new DropTargetAdapter() {
            @Override
            public void drop(DropTargetEvent event) {
                if (event.data instanceof String[] files && files.length > 0) {
                    loadMsgFile(new File(files[0]));
                }
            }
        });
    }

    private void openFileDialog() {
        FileDialog fd = new FileDialog(shell, SWT.OPEN);
        fd.setText("MSGファイルを開く");
        fd.setFilterExtensions(new String[]{"*.msg", "*.*"});
        fd.setFilterNames(new String[]{"Outlook MSG (*.msg)", "すべてのファイル (*.*)"});
        String path = fd.open();
        if (path != null) {
            loadMsgFile(new File(path));
        }
    }

    private void loadMsgFile(File file) {
        statusBar.setText("読み込み中: " + file.getName());
        shell.setCursor(display.getSystemCursor(SWT.CURSOR_WAIT));

        display.asyncExec(() -> {
            try {
                currentMsgFileBytes = Files.readAllBytes(file.toPath());
                currentMsgHash = CommentRepository.hashBytes(currentMsgFileBytes);
                currentMsg = parser.parse(file);

                // Update mail detail
                mailDetailBrowser.setText(buildMailDetailHtml(currentMsg));

                // Update attachment tree
                attachmentTree.removeAll();
                if (currentMsg.getAttachments() != null) {
                    for (AttachmentData att : currentMsg.getAttachments()) {
                        TreeItem item = new TreeItem(attachmentTree, SWT.NONE);
                        item.setText(att.getFileName() + " (" + att.getFormattedSize() + ")");
                        item.setData(att);
                    }
                }

                // Load comments for the MSG itself
                currentAttachmentHash = currentMsgHash;
                loadComments(currentMsgHash);

                statusBar.setText("読み込み完了: " + file.getName()
                        + " — 添付ファイル: " + (currentMsg.getAttachments() != null ? currentMsg.getAttachments().size() : 0) + "件");
                shell.setText("MSG Viewer — " + (currentMsg.getSubject() != null ? currentMsg.getSubject() : file.getName()));

            } catch (Exception e) {
                log.error("MSG load error", e);
                MessageBox mb = new MessageBox(shell, SWT.ICON_ERROR | SWT.OK);
                mb.setText("エラー");
                mb.setMessage("MSGファイルの読み込みに失敗しました:\n" + e.getMessage());
                mb.open();
                statusBar.setText("エラー: " + e.getMessage());
            } finally {
                shell.setCursor(null);
            }
        });
    }

    private void onAttachmentSelected() {
        TreeItem[] sel = attachmentTree.getSelection();
        if (sel.length == 0) return;

        selectedAttachment = (AttachmentData) sel[0].getData();
        if (selectedAttachment == null) return;

        // Update comment context
        currentAttachmentHash = CommentRepository.hashBytes(selectedAttachment.getData());
        loadComments(currentAttachmentHash);

        String name = selectedAttachment.getFileName().toLowerCase();
        String mime = selectedAttachment.getMimeType() != null ? selectedAttachment.getMimeType().toLowerCase() : "";

        // Handle ZIP separately
        if (name.endsWith(".zip") || mime.equals("application/zip")) {
            handleZipAttachment(selectedAttachment);
            return;
        }

        // Generate preview
        statusBar.setText("プレビュー生成中: " + selectedAttachment.getFileName());
        display.asyncExec(() -> {
            try {
                PreviewResult result = previewEngine.preview(selectedAttachment);
                showPreview(result);
                statusBar.setText("プレビュー: " + selectedAttachment.getFileName());
            } catch (Exception e) {
                log.error("Preview error", e);
                statusBar.setText("プレビューエラー: " + e.getMessage());
            }
        });
    }

    private void handleZipAttachment(AttachmentData att) {
        try {
            boolean encrypted = zipHandler.isEncrypted(att.getData());
            char[] password = null;

            if (encrypted) {
                InputDialog dlg = new InputDialog(shell, "パスワード入力",
                        "このZIPファイルはパスワードで保護されています。\nパスワードを入力してください:");
                String pw = dlg.open();
                if (pw == null) return; // cancelled
                password = pw.toCharArray();
            }

            // Extract and show tree of contents
            java.nio.file.Path extractDir = zipHandler.extract(att.getData(), password);
            // Show extracted files in a message or tree (simplified)
            StringBuilder html = new StringBuilder();
            html.append("<html><head><meta charset='UTF-8'><style>")
                .append("body{font-family:sans-serif;padding:15px;}")
                .append("li{padding:3px 0;}")
                .append("</style></head><body>")
                .append("<h3>📦 ZIP展開結果</h3><ul>");

            Files.walk(extractDir).filter(Files::isRegularFile).forEach(p -> {
                html.append("<li>").append(extractDir.relativize(p).toString()).append("</li>");
            });

            html.append("</ul></body></html>");
            showHtmlPreview(html.toString());
            statusBar.setText("ZIP展開完了: " + att.getFileName());

            // Schedule cleanup
            Runtime.getRuntime().addShutdownHook(new Thread(() -> ZipHandler.deleteDirectory(extractDir)));

        } catch (Exception e) {
            log.error("ZIP handling error", e);
            MessageBox mb = new MessageBox(shell, SWT.ICON_ERROR | SWT.OK);
            mb.setText("ZIPエラー");
            mb.setMessage("ZIPファイルの処理に失敗しました:\n" + e.getMessage());
            mb.open();
        }
    }

    private void showPreview(PreviewResult result) {
        StackLayout layout = (StackLayout) previewContainer.getLayout();

        switch (result.getType()) {
            case HTML -> {
                previewBrowser.setText(result.getContent());
                layout.topControl = previewBrowser;
            }
            case TEXT -> {
                previewBrowser.setText("<html><head><meta charset='UTF-8'><style>"
                        + "body{font-family:monospace;padding:15px;white-space:pre-wrap;}"
                        + "</style></head><body>"
                        + escapeHtml(result.getContent())
                        + "</body></html>");
                layout.topControl = previewBrowser;
            }
            case IMAGE -> {
                showImagePreview(result.getImageData());
                layout.topControl = previewCanvas;
                // Handle paging
                currentPage = result.getCurrentPage();
                totalPages = result.getPageCount();
                if (totalPages > 1) {
                    statusBar.setText(String.format("ページ %d / %d — 左右キーでページ送り", currentPage + 1, totalPages));
                }
            }
            case UNSUPPORTED -> {
                String html = "<html><head><meta charset='UTF-8'><style>"
                        + "body{font-family:sans-serif;padding:20px;}"
                        + ".info{background:#f0f4f8;padding:20px;border-radius:8px;}"
                        + "</style></head><body><div class='info'>"
                        + "<h3>📄 " + escapeHtml(result.getFileName()) + "</h3>"
                        + "<p>MIME: " + escapeHtml(result.getMimeType()) + "</p>"
                        + "<p>このファイル形式のプレビューには対応していません。<br>"
                        + "右クリックメニューからファイルを保存できます。</p>"
                        + "</div></body></html>";
                previewBrowser.setText(html);
                layout.topControl = previewBrowser;
            }
        }
        previewContainer.layout();
    }

    private void showHtmlPreview(String html) {
        StackLayout layout = (StackLayout) previewContainer.getLayout();
        previewBrowser.setText(html);
        layout.topControl = previewBrowser;
        previewContainer.layout();
    }

    private void showImagePreview(byte[] imageData) {
        // Dispose old image listener
        Listener[] listeners = previewCanvas.getListeners(SWT.Paint);
        for (Listener l : listeners) previewCanvas.removeListener(SWT.Paint, l);

        try {
            Image img = new Image(display, new ByteArrayInputStream(imageData));
            previewCanvas.addListener(SWT.Paint, event -> {
                Rectangle bounds = img.getBounds();
                Rectangle clientArea = previewCanvas.getClientArea();
                // Scale to fit
                double scaleX = (double) clientArea.width / bounds.width;
                double scaleY = (double) clientArea.height / bounds.height;
                double scale = Math.min(scaleX, scaleY);
                int w = (int) (bounds.width * scale);
                int h = (int) (bounds.height * scale);
                int x = (clientArea.width - w) / 2;
                int y = (clientArea.height - h) / 2;
                event.gc.drawImage(img, 0, 0, bounds.width, bounds.height, x, y, w, h);
            });
            previewCanvas.addListener(SWT.Dispose, e -> img.dispose());

            // Add keyboard page navigation for PDF/PPT
            previewCanvas.addListener(SWT.KeyDown, event -> {
                if (totalPages > 1 && selectedAttachment != null) {
                    int newPage = currentPage;
                    if (event.keyCode == SWT.ARROW_RIGHT || event.keyCode == SWT.PAGE_DOWN) {
                        newPage = Math.min(currentPage + 1, totalPages - 1);
                    } else if (event.keyCode == SWT.ARROW_LEFT || event.keyCode == SWT.PAGE_UP) {
                        newPage = Math.max(currentPage - 1, 0);
                    }
                    if (newPage != currentPage) {
                        final int page = newPage;
                        display.asyncExec(() -> {
                            PreviewResult result = previewEngine.previewPage(selectedAttachment, page);
                            if (result.getType() == PreviewResult.Type.IMAGE) {
                                showPreview(result);
                            }
                        });
                    }
                }
            });
            previewCanvas.setFocus();
            previewCanvas.redraw();
        } catch (Exception e) {
            log.error("Image preview error", e);
        }
    }

    // ===================== Comments =====================

    private void loadComments(String fileHash) {
        commentList.removeAll();
        List<Comment> comments = commentRepo.findByFileHash(fileHash);
        for (Comment c : comments) {
            String display = "[" + c.getCreatedAt().format(DT_FMT) + "] " + c.getText();
            commentList.add(display);
            commentList.setData(String.valueOf(commentList.getItemCount() - 1), c);
        }
    }

    private void addComment() {
        String text = commentInput.getText().trim();
        if (text.isEmpty() || currentAttachmentHash == null) return;

        String fileName = (selectedAttachment != null) ? selectedAttachment.getFileName()
                : (currentMsg != null ? currentMsg.getSubject() : "unknown");

        Comment c = new Comment(currentAttachmentHash, fileName, text);
        commentRepo.create(c);
        commentInput.setText("");
        loadComments(currentAttachmentHash);
    }

    private void deleteSelectedComment() {
        int idx = commentList.getSelectionIndex();
        if (idx < 0) return;

        Comment c = (Comment) commentList.getData(String.valueOf(idx));
        if (c == null) return;

        MessageBox confirm = new MessageBox(shell, SWT.ICON_QUESTION | SWT.YES | SWT.NO);
        confirm.setText("確認");
        confirm.setMessage("このコメントを削除しますか？");
        if (confirm.open() == SWT.YES) {
            commentRepo.delete(c.getId());
            loadComments(currentAttachmentHash);
        }
    }

    // ===================== HTML Helpers =====================

    private String buildMailDetailHtml(MsgData msg) {
        StringBuilder html = new StringBuilder();
        html.append("<html><head><meta charset='UTF-8'><style>")
            .append("body{font-family:sans-serif;padding:15px;margin:0;background:#fafafa;}")
            .append(".header{background:#fff;padding:15px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.1);margin-bottom:15px;}")
            .append(".field{margin:4px 0;}.label{font-weight:bold;color:#555;display:inline-block;width:70px;}")
            .append(".subject{font-size:1.3em;font-weight:bold;color:#2c3e50;margin-bottom:10px;}")
            .append(".body-content{background:#fff;padding:15px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.1);}")
            .append("</style></head><body>");

        html.append("<div class='header'>");
        html.append("<div class='subject'>").append(escapeHtml(msg.getSubject())).append("</div>");
        html.append("<div class='field'><span class='label'>差出人:</span> ").append(escapeHtml(msg.getFrom())).append("</div>");
        html.append("<div class='field'><span class='label'>宛先:</span> ").append(escapeHtml(msg.getTo())).append("</div>");
        if (msg.getCc() != null && !msg.getCc().isEmpty()) {
            html.append("<div class='field'><span class='label'>CC:</span> ").append(escapeHtml(msg.getCc())).append("</div>");
        }
        html.append("<div class='field'><span class='label'>日時:</span> ").append(escapeHtml(msg.getDate())).append("</div>");
        html.append("</div>");

        html.append("<div class='body-content'>");
        if (msg.getBodyHtml() != null && !msg.getBodyHtml().isEmpty()) {
            html.append(msg.getBodyHtml());
        } else if (msg.getBodyText() != null && !msg.getBodyText().isEmpty()) {
            html.append("<pre style='white-space:pre-wrap;'>").append(escapeHtml(msg.getBodyText())).append("</pre>");
        } else {
            html.append("<p style='color:#999;'>(本文なし)</p>");
        }
        html.append("</div></body></html>");

        return html.toString();
    }

    private String welcomeHtml() {
        return "<html><head><meta charset='UTF-8'><style>"
                + "body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f7fa;}"
                + ".welcome{text-align:center;color:#888;}"
                + ".welcome h1{font-size:2em;color:#4a90d9;}"
                + ".welcome p{font-size:1.1em;}"
                + "</style></head><body><div class='welcome'>"
                + "<h1>📧 MSG Viewer</h1>"
                + "<p>MSGファイルをドラッグ&ドロップするか、<br>「ファイル→開く」で選択してください</p>"
                + "</div></body></html>";
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }

    // ===================== Input Dialog =====================

    /**
     * Simple password input dialog.
     */
    private static class InputDialog {
        private final Shell parent;
        private final String title;
        private final String message;
        private String result;

        InputDialog(Shell parent, String title, String message) {
            this.parent = parent;
            this.title = title;
            this.message = message;
        }

        String open() {
            Shell dialog = new Shell(parent, SWT.DIALOG_TRIM | SWT.APPLICATION_MODAL);
            dialog.setText(title);
            dialog.setLayout(new GridLayout(2, false));

            Label msgLabel = new Label(dialog, SWT.WRAP);
            GridData msgData = new GridData(SWT.FILL, SWT.CENTER, true, false, 2, 1);
            msgData.widthHint = 300;
            msgLabel.setLayoutData(msgData);
            msgLabel.setText(message);

            Text input = new Text(dialog, SWT.BORDER | SWT.PASSWORD);
            input.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 2, 1));

            Button okBtn = new Button(dialog, SWT.PUSH);
            okBtn.setText("OK");
            okBtn.setLayoutData(new GridData(SWT.RIGHT, SWT.CENTER, true, false));
            okBtn.addSelectionListener(new SelectionAdapter() {
                @Override
                public void widgetSelected(SelectionEvent e) {
                    result = input.getText();
                    dialog.close();
                }
            });

            Button cancelBtn = new Button(dialog, SWT.PUSH);
            cancelBtn.setText("キャンセル");
            cancelBtn.addSelectionListener(new SelectionAdapter() {
                @Override
                public void widgetSelected(SelectionEvent e) {
                    result = null;
                    dialog.close();
                }
            });

            dialog.setDefaultButton(okBtn);
            dialog.pack();
            // Center on parent
            Rectangle parentBounds = parent.getBounds();
            Point size = dialog.getSize();
            dialog.setLocation(parentBounds.x + (parentBounds.width - size.x) / 2,
                    parentBounds.y + (parentBounds.height - size.y) / 2);
            dialog.open();

            Display display = parent.getDisplay();
            while (!dialog.isDisposed()) {
                if (!display.readAndDispatch()) display.sleep();
            }
            return result;
        }
    }
}
