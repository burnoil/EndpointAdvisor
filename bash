exists key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry 
AND 
exists value "ProductReleaseIds" whose(
    (it as string contains "ProPlusVolume") OR 
    (it as string contains "ProPlus2019Volume") OR 
    (it as string contains "ProPlus2021Volume") OR 
    (it as string contains "ProPlus2024Volume")
) of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
```

**Note:** Office 2016 Volume is just `ProPlusVolume` (without the year).

## If You Have Retail Editions Too
```
exists key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry 
AND 
exists value "ProductReleaseIds" whose(
    (it as string contains "ProPlusVolume") OR 
    (it as string contains "ProPlusRetail") OR
    (it as string contains "ProPlus2019Volume") OR 
    (it as string contains "ProPlus2019Retail") OR
    (it as string contains "ProPlus2021Volume") OR 
    (it as string contains "ProPlus2021Retail") OR
    (it as string contains "ProPlus2024Volume") OR
    (it as string contains "ProPlus2024Retail")
) of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry
```

## Updated Analysis Property

Update Property 4 to include 2016:

### Property 4: Needs M365 Upgrade (2016/2019/2021/2024)
```
exists value "ProductReleaseIds" whose (
    (it as string contains "ProPlusVolume") OR 
    (it as string contains "ProPlusRetail") OR
    (it as string contains "ProPlus2019") OR 
    (it as string contains "ProPlus2021") OR 
    (it as string contains "ProPlus2024")
) of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry AND NOT (exists value "ProductReleaseIds" whose (it as string contains "O365ProPlusRetail") of key "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" of native registry)
