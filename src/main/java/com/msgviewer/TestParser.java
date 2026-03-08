package com.msgviewer;

import java.nio.file.Path;
import java.util.List;

/**
 * Integration test: parses downloaded MSG files, runs preview engine,
 * and exercises comment CRUD — all without launching the GUI.
 */
public class TestParser {

    public static void main(String[] args) throws Exception {
        Path sampleDir = Path.of("sample-msg");
        List<String> fileNames = List.of(
                "simple_test_msg.msg",
                "quick.msg",
                "attachment_test_msg.msg"
        );

        MsgParser parser = new MsgParser();
        PreviewEngine preview = new PreviewEngine();
        int testNum = 0;

        System.out.println("========================================");
        System.out.println(" MSG Viewer 統合テスト");
        System.out.println("========================================\n");

        // --- Test 1-3: MSG Parse + Preview ---
        for (String fn : fileNames) {
            testNum++;
            Path p = sampleDir.resolve(fn);
            System.out.println("--- テスト " + testNum + ": " + fn + " の解析 ---");
            try {
                MsgData data = parser.parse(p.toFile());
                System.out.println("  件名     : " + data.getSubject());
                System.out.println("  差出人   : " + data.getFrom());
                System.out.println("  宛先     : " + data.getTo());
                System.out.println("  CC       : " + data.getCc());
                System.out.println("  日時     : " + data.getDate());
                String body = data.getBodyText();
                System.out.println("  本文長   : " + (body != null ? body.length() : 0) + " 文字");
                if (body != null && !body.isBlank()) {
                    System.out.println("  本文先頭 : " + body.substring(0, Math.min(80, body.length())).replace("\n", "\\n"));
                }
                System.out.println("  HTML本文 : " + (data.getBodyHtml() != null && !data.getBodyHtml().isBlank() ? "あり" : "なし"));
                System.out.println("  添付数   : " + data.getAttachments().size());

                for (int i = 0; i < data.getAttachments().size(); i++) {
                    AttachmentData att = data.getAttachments().get(i);
                    System.out.println("    添付[" + i + "] : " + att.getFileName()
                            + " (" + att.getMimeType() + ", " + att.getFormattedSize() + ")");

                    // Preview each attachment
                    if (att.getData() != null) {
                        try {
                            PreviewResult result = preview.preview(att);
                            System.out.println("    プレビュー: type=" + result.getType()
                                    + ", pages=" + result.getPageCount());
                        } catch (Exception pe) {
                            System.out.println("    プレビュー: エラー - " + pe.getMessage());
                        }
                    }
                }
                System.out.println("  ✅ 解析成功");
            } catch (Exception e) {
                System.out.println("  ❌ 解析失敗: " + e.getMessage());
            }
            System.out.println();
        }

        // --- Test 4: Comment CRUD ---
        testNum++;
        System.out.println("--- テスト " + testNum + ": コメント CRUD (H2 Database) ---");
        try {
            CommentRepository repo = new CommentRepository();

            // Create
            String fileHash = "test-hash-12345";
            Comment c1 = new Comment(fileHash, "test.msg", "これはテストコメントです。");
            repo.create(c1);
            Comment c2 = new Comment(fileHash, "test.msg", "2番目のコメント。");
            repo.create(c2);

            // Read
            List<Comment> comments = repo.findByFileHash(fileHash);
            System.out.println("  作成後コメント数: " + comments.size());
            for (Comment c : comments) {
                System.out.println("    [" + c.getCreatedAt() + "] " + c.getText());
            }

            // Delete
            if (!comments.isEmpty()) {
                repo.delete(comments.get(0).getId());
                List<Comment> after = repo.findByFileHash(fileHash);
                System.out.println("  削除後コメント数: " + after.size());
            }

            // Clean up test data
            for (Comment c : repo.findByFileHash(fileHash)) {
                repo.delete(c.getId());
            }

            System.out.println("  ✅ コメント CRUD 成功");
        } catch (Exception e) {
            System.out.println("  ❌ コメント CRUD 失敗: " + e.getMessage());
            e.printStackTrace();
        }
        System.out.println();

        System.out.println("========================================");
        System.out.println(" テスト完了: " + testNum + " テスト実行");
        System.out.println("========================================");
    }
}
