///usr/bin/env jbang "$0" "$@" ; exit $?
//JAVA 11+
//COMPILE_OPTIONS -encoding UTF-8
//#JAVAOPTIONS -XstartOnFirstThread
//DEPS org.eclipse.platform:org.eclipse.swt.cocoa.macosx.x86_64:3.127.0
//#DEPS org.eclipse.platform:org.eclipse.swt.cocoa.macosx.aarch64:3.127.0
//#DEPS org.eclipse.platform:org.eclipse.swt.win32.win32.x86_64:3.127.0
//#DEPS org.eclipse.platform:org.eclipse.org.eclipse.swt.gtk.linux.x86_64:3.127.0
//#DEPS org.eclipse.platform:org.eclipse.org.eclipse.swt.gtk.linux.aarch64:3.127.0
//DEPS org.apache.poi:poi:5.3.0
//DEPS org.apache.poi:poi-scratchpad:5.3.0
//DEPS org.apache.poi:poi-ooxml:5.3.0
//DEPS org.apache.pdfbox:pdfbox:3.0.3
//DEPS org.apache.tika:tika-core:2.9.2
//DEPS net.lingala.zip4j:zip4j:2.11.5
//DEPS com.h2database:h2:2.3.232
//DEPS org.slf4j:slf4j-api:2.0.16
//DEPS org.slf4j:slf4j-simple:2.0.16
//DEPS com.github.albfernandez:juniversalchardet:2.5.0
//DEPS jakarta.mail:jakarta.mail-api:2.1.3
//DEPS org.eclipse.angus:angus-mail:2.0.3
//SOURCES ./AttachmentData.java
//SOURCES ./Comment.java
//SOURCES ./CommentRepository.java
//SOURCES ./MainWindow.java
//SOURCES ./MsgParser.java
//SOURCES ./MsgData.java
//SOURCES ./PreviewEngine.java
//SOURCES ./PreviewResult.java
//SOURCES ./ZipHandler.java
//MAIN com.msgviewer.App

package com.msgviewer;

import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.nio.file.*;

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
