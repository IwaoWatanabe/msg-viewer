
## Windows で起動するときの注意

App.java のSWTのArtifact-IDを Windows用に調整してください。

Artifact-ID | 説明
--------------|------
org.eclipse.swt.cocoa.win32.win32 | Intelプロセッサ向け
org.eclipse.swt.cocoa.win32.aarch64 | ARM64プロセッサ向け

## Linuxもしくは　ChromOSのLinuxエミュレータもしくは　Windows のWSLで起動するときの注意

App.java の先頭部に定義されている
SWTのArtifact-IDを Linux用に調整してください。

Artifact-ID | 説明
--------------|------
org.eclipse.swt.gtk.linux.x86_64 | Intelプロセッサ向け
org.eclipse.swt.gtk.linux.aarch64 | ARM64プロセッサ向け

## macOSで起動するときの注意

App.java の先頭部に定義されている
SWTのArtifact-IDを macOS用に調整してください。

Artifact-ID | 説明
--------------|------
org.eclipse.swt.cocoa.macosx.x86_64 | Intelプロセッサ向け
org.eclipse.swt.cocoa.macosx.aarch64 | Mx プロセッサ向け

また起動時に次のようにVMオプションを指定する必要があります。

```
jbang run --java-options="-XstartOnFirstThread" App.java
```

