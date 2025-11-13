At C:\Program Files\LLEA\download_LLEA.ps1:16 char:25
+ foreach ($item in $urls)
+                         ~
Missing statement body in foreach loop.
At C:\Program Files\LLEA\download_LLEA.ps1:18 char:20
+     $url = $item[0]
+                    ~
Missing closing ')' in expression.
At C:\Program Files\LLEA\download_LLEA.ps1:19 char:5
+     $dest = $item[1]
+     ~~~~~
Unexpected token '$dest' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:22 char:33
+     for ($i = 1; $i -le 3; $i++)
+                                 ~
Missing statement body in for loop.
At C:\Program Files\LLEA\download_LLEA.ps1:24 char:12
+         try
+            ~
Missing closing ')' in expression.
At C:\Program Files\LLEA\download_LLEA.ps1:25 char:9
+         (
+         ~
Unexpected token '(' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:26 char:79
+ ...    Write-Host "Downloading $(Split-Path $dest -Leaf) (attempt $i)..."
+                                                                          ~
Missing closing ')' in expression.
At C:\Program Files\LLEA\download_LLEA.ps1:27 char:13
+             Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsi ...
+             ~~~~~~~~~~~~~~~~~
Unexpected token 'Invoke-WebRequest' in expression or statement.
At C:\Program Files\LLEA\download_LLEA.ps1:29 char:32
+             if (Test-Path $dest)
+                                ~
Missing statement block after if ( condition ).
At C:\Program Files\LLEA\download_LLEA.ps1:31 char:52
+                 $fileSize = (Get-Item $dest).Length
+                                                    ~
Missing closing ')' in expression.
Not all parse errors were reported.  Correct the reported errors and try again.
    + CategoryInfo          : ParserError: (:) [], ParseException
    + FullyQualifiedErrorId : MissingForeachStatement
