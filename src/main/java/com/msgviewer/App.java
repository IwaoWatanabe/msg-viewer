package com.msgviewer;

import com.msgviewer.ui.MainWindow;
import com.msgviewer.preview.ZipHandler;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;

/**
 * MSG Viewer Application entry point.
 * Launches the SWT desktop GUI for viewing Outlook MSG files.
 */
public class App {

    private static final Logger log = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) {
        // Register shutdown hook to clean up temp files
        registerCleanupHook();

        Display display = new Display();
        try {
            Shell shell = new Shell(display);
            shell.setText("MSG Viewer");
            shell.setSize(1200, 800);

            // Initialize main window (creates all UI components)
            new MainWindow(shell);

            shell.open();
            while (!shell.isDisposed()) {
                if (!display.readAndDispatch()) {
                    display.sleep();
                }
            }
        } catch (Exception e) {
            log.error("Application error", e);
        } finally {
            display.dispose();
        }
    }

    private static void registerCleanupHook() {
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            log.info("Cleaning up temporary files...");
            try {
                Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"));
                if (Files.exists(tempDir)) {
                    Files.list(tempDir)
                        .filter(p -> p.getFileName().toString().startsWith("msg-viewer-"))
                        .forEach(p -> ZipHandler.deleteDirectory(p));
                }
            } catch (Exception e) {
                // Shutdown, ignore errors
            }
        }));
    }
}
