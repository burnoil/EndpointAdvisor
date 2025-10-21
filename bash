<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <Analysis>
    <Title>Microsoft Office Version Inventory (C2R + MSI + Project/Visio)</Title>
    <Description><![CDATA[
Reports Microsoft Office inventory including Click-to-Run (Microsoft 365 Apps / Office 2019/2021) and legacy MSI installs, plus explicit detection of Project and Visio (C2R and MSI).

Properties:
- Office (C2R): ProductReleaseIds, Version, Channel, Architecture, InstallPath
- Office (MSI): Discovered suites
- Project/Visio: C2R present?, C2R release IDs, MSI versions

Activation Relevance: Windows only
]]></Description>
    <Relevance>(name of operating system as lowercase starts with "win")</Relevance>
    <Category>Inventory</Category>
    <Source>Timerider</Source>
    <SourceReleaseDate>2025-10-21</SourceReleaseDate>
    <DefaultFlag>true</DefaultFlag>
    <Properties>
      
      <!-- Office (C2R) core -->
      <Property Name="Office (C2R) ProductReleaseIds" ID="1">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "ProductReleaseIds" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "N/A"
        </Relevance>
      </Property>

      <Property Name="Office (C2R) Version" ID="2">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "ClientVersionToReport" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "N/A"
        </Relevance>
      </Property>

      <Property Name="Office (C2R) Channel" ID="3">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "UpdateChannel" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "N/A"
        </Relevance>
      </Property>

      <Property Name="Office (C2R) Architecture" ID="4">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "Platform" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "Unknown"
        </Relevance>
      </Property>

      <Property Name="Office (C2R) InstallPath" ID="5">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "InstallPath" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "N/A"
        </Relevance>
      </Property>

      <!-- Office (MSI) suites -->
      <Property Name="Office (MSI) Suites" ID="6">
        <Relevance>
(if exists keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" of native registry then
  ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" of native registry )
else
  (nothing)) ;
(if exists keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry then
  ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry )
else
  (nothing))
        </Relevance>
      </Property>

      <!-- Visio (C2R / MSI) -->
      <Property Name="Visio (C2R) Present" ID="7">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (exists matches (regex "(?i)\bVisio\w*\b") of (value "ProductReleaseIds" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string))
else false
        </Relevance>
      </Property>

      <Property Name="Visio (C2R) ReleaseIds" ID="8">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  ( (it as string) of (substrings separated by "," of (value "ProductReleaseIds" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string) whose (it as lowercase contains "visio")) )
else "N/A"
        </Relevance>
      </Property>

      <Property Name="Visio (MSI) Version(s)" ID="9">
        <Relevance>
( ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase contains "visio")) of key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" of native registry )
;
  ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase contains "visio")) of key "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry ) )
        </Relevance>
      </Property>

      <!-- Project (C2R / MSI) -->
      <Property Name="Project (C2R) Present" ID="10">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (exists matches (regex "(?i)\bProject\w*\b") of (value "ProductReleaseIds" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string))
else false
        </Relevance>
      </Property>

      <Property Name="Project (C2R) ReleaseIds" ID="11">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  ( (it as string) of (substrings separated by "," of (value "ProductReleaseIds" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string) whose (it as lowercase contains "project")) )
else "N/A"
        </Relevance>
      </Property>

      <Property Name="Project (MSI) Version(s)" ID="12">
        <Relevance>
( ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase contains "project")) of key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" of native registry )
;
  ( (value "DisplayName" of it as string & " - " & (if exists value "DisplayVersion" of it then value "DisplayVersion" of it as string else "Unknown")) of keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase contains "project")) of key "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry ) )
        </Relevance>
      </Property>

      <!-- Helpful rollups -->
      <Property Name="Office Installed Type (C2R/MSI/None)" ID="13">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then "C2R"
else if (exists keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" of native registry or
         exists keys whose (exists value "DisplayName" of it and (value "DisplayName" of it as string as lowercase starts with "microsoft office")) of key "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" of native registry) then "MSI"
else "None"
        </Relevance>
      </Property>

      <Property Name="Office (C2R) CDNBaseUrl (if present)" ID="14">
        <Relevance>
if exists key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry and exists value "CDNBaseUrl" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry then
  (value "CDNBaseUrl" of key "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" of native registry as string)
else "N/A"
        </Relevance>
      </Property>

    </Properties>
  </Analysis>
</BES>
