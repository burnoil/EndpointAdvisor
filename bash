At C:\Program Files\LLEA\download_LLEA.ps1:18 char:99
+ ... { $ok = $true; Write-Host "  OK: $((Get-Item $dest).Length) bytes" }}
+                                                                          ~
The Try statement is missing its Catch or Finally block.
At C:\Program Files\LLEA\download_LLEA.ps1:23 char:5
+     }}
+     ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:23 char:6
+     }}
+      ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:24 char:3
+   }}
+   ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:24 char:4
+   }}
+    ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:25 char:54
+   if (-not $ok) { $failed += Split-Path $dest -Leaf }}
+                                                      ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:26 char:1
+ }}
+ ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:26 char:2
+ }}
+  ~
Unexpected token '}' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:28 char:78
+ ... ed.Count -gt 0) { Write-Host "ERROR: $($failed -join ',')"; exit 1 }}
+                                                                         ~
Unexpected token '}' in expression or statement.
    + CategoryInfo          : ParserError: (:) [], ParseException
    + FullyQualifiedErrorId : MissingCatchOrFinally
