package com.msgviewer;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.MessageDigest;
import java.sql.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * DAO for managing comments stored in H2 embedded database.
 * Database file is stored at ~/.msg-viewer/comments.mv.db
 */
class CommentRepository {

    private static final Logger log = LoggerFactory.getLogger(CommentRepository.class);
    private final String dbUrl;

    public CommentRepository() {
        Path dbDir = Paths.get(System.getProperty("user.home"), ".msg-viewer");
        try {
            Files.createDirectories(dbDir);
        } catch (Exception e) {
            log.error("DB directory creation failed", e);
        }
        Path dbFile = dbDir.resolve("comments");
        this.dbUrl = "jdbc:h2:" + dbFile.toAbsolutePath();
        initializeSchema();
    }

    private void initializeSchema() {
        try (Connection conn = getConnection();
             Statement stmt = conn.createStatement()) {
            stmt.execute("""
                CREATE TABLE IF NOT EXISTS comments (
                    id BIGINT AUTO_INCREMENT PRIMARY KEY,
                    file_hash VARCHAR(128) NOT NULL,
                    file_name VARCHAR(512),
                    text CLOB NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            """);
            stmt.execute("""
                CREATE INDEX IF NOT EXISTS idx_comments_file_hash ON comments(file_hash)
            """);
            log.info("H2 Database initialized: {}", dbUrl);
        } catch (SQLException e) {
            log.error("Schema initialization failed", e);
            throw new RuntimeException("データベースの初期化に失敗しました", e);
        }
    }

    private Connection getConnection() throws SQLException {
        return DriverManager.getConnection(dbUrl, "sa", "");
    }

    /**
     * Creates a new comment.
     */
    public Comment create(Comment comment) {
        String sql = "INSERT INTO comments (file_hash, file_name, text, created_at, updated_at) VALUES (?, ?, ?, ?, ?)";
        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            LocalDateTime now = LocalDateTime.now();
            comment.setCreatedAt(now);
            comment.setUpdatedAt(now);
            ps.setString(1, comment.getFileHash());
            ps.setString(2, comment.getFileName());
            ps.setString(3, comment.getText());
            ps.setTimestamp(4, Timestamp.valueOf(now));
            ps.setTimestamp(5, Timestamp.valueOf(now));
            ps.executeUpdate();

            try (ResultSet rs = ps.getGeneratedKeys()) {
                if (rs.next()) {
                    comment.setId(rs.getLong(1));
                }
            }
        } catch (SQLException e) {
            log.error("Comment creation failed", e);
        }
        return comment;
    }

    /**
     * Finds all comments for a given file hash.
     */
    public List<Comment> findByFileHash(String fileHash) {
        List<Comment> comments = new ArrayList<>();
        String sql = "SELECT * FROM comments WHERE file_hash = ? ORDER BY created_at DESC";
        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, fileHash);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    comments.add(mapRow(rs));
                }
            }
        } catch (SQLException e) {
            log.error("Comment query failed", e);
        }
        return comments;
    }

    /**
     * Updates the text of an existing comment.
     */
    public void update(long id, String newText) {
        String sql = "UPDATE comments SET text = ?, updated_at = ? WHERE id = ?";
        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setString(1, newText);
            ps.setTimestamp(2, Timestamp.valueOf(LocalDateTime.now()));
            ps.setLong(3, id);
            ps.executeUpdate();
        } catch (SQLException e) {
            log.error("Comment update failed", e);
        }
    }

    /**
     * Deletes a comment by ID.
     */
    public void delete(long id) {
        String sql = "DELETE FROM comments WHERE id = ?";
        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setLong(1, id);
            ps.executeUpdate();
        } catch (SQLException e) {
            log.error("Comment deletion failed", e);
        }
    }

    private Comment mapRow(ResultSet rs) throws SQLException {
        Comment c = new Comment();
        c.setId(rs.getLong("id"));
        c.setFileHash(rs.getString("file_hash"));
        c.setFileName(rs.getString("file_name"));
        c.setText(rs.getString("text"));
        c.setCreatedAt(rs.getTimestamp("created_at").toLocalDateTime());
        c.setUpdatedAt(rs.getTimestamp("updated_at").toLocalDateTime());
        return c;
    }

    /**
     * Utility: generates a SHA-256 hash string for file identification.
     */
    public static String hashBytes(byte[] data) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] digest = md.digest(data);
            StringBuilder sb = new StringBuilder();
            for (byte b : digest) {
                sb.append(String.format("%02x", b));
            }
            return sb.toString();
        } catch (Exception e) {
            return String.valueOf(java.util.Arrays.hashCode(data));
        }
    }
}
