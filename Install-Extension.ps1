[CmdletBinding(DefaultParametersetName = 'Install')] 
param( 
    # Use UpdateSearchString or UpdateID and RevisionNumber to find the Update
    [Parameter(Position = 0, Mandatory = $true, ParameterSetName = 'Install')] [string]$SQLServer, # e.g. = sqlserver.domain.com
    [Parameter(Position = 1, Mandatory = $true, ParameterSetName = 'Install')] [string]$SQLDBName, # e.g. = SUSDB
    [Parameter(Mandatory = $false, ParameterSetName = 'Uninstall')] [switch]$Uninstall
    
)
if ($PsCmdLet.ParameterSetName -eq "Uninstall") {
    Get-ChildItem -Path "$($ENV:SMS_ADMIN_UI_PATH)\..\..\XmlStorage\Extensions\Actions" -Recurse -Filter "ShowApplicabilityRule*.xml" | Remove-Item 
    Get-ChildItem -Path "$($ENV:SMS_ADMIN_UI_PATH)\..\..\XmlStorage\Tools" -Recurse -Filter "Show-ApplicabilityRule.ps1" | Remove-Item
}
else {
    $Extension = @"
<ActionDescription Class="Executable" DisplayName="Show Applicability Rule" MnemonicDisplayName="Show Applicability Rule" Description="Displays the applicability rule for a selected update" SqmDataPoint="53">
  <ShowOn>
    <string>DefaultHomeTab</string>
    <string>ContextMenu</string>
  </ShowOn>
 <ResourceAssembly>
    <Assembly>AdminUI.SoftwareUpdateProperties.dll</Assembly>
    <Type>Microsoft.ConfigurationManagement.AdminConsole.SoftwareUpdateProperties.Properties.Resources.resources</Type>
  </ResourceAssembly>
  <ImagesDescription>
    <ResourceAssembly>
      <Assembly>AdminUI.UIResources.dll</Assembly>
      <Type>Microsoft.ConfigurationManagement.AdminConsole.UIResources.Properties.Resources.resources</Type>
    </ResourceAssembly>
    <ImageResourceName>SUM_Update</ImageResourceName>
  </ImagesDescription>
  <Executable>
    <FilePath>"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"</FilePath>
    <Parameters>-ExecutionPolicy ByPass -File "#PATHTOSCRIPT#" -SQLServer "#SQLSERVER#" -SQLDBName "#SUSDB#" #PARAMS# </Parameters>
  </Executable>
</ActionDescription>
"@
    Push-Location $PSScriptRoot
    #Create Tools Folder and copy Show-ApplicabilityRule.ps1 into it
    $ToolsFolder = New-Item -Path "$($ENV:SMS_ADMIN_UI_PATH)\..\..\XmlStorage\" -Name "Tools" -ItemType "directory" -Force
    Copy-Item .\Show-ApplicabilityRule.ps1 $ToolsFolder 
    $PathToScript = "$ToolsFolder\Show-ApplicabilityRule.ps1" 

    #Create GUID Folder and copy ShowApplicabilityRule.xml
    $Guids = ("Office_365_Updates", "ef93deea-a0bc-4f36-9b48-a510ac5340eb"),
    ("All_Software_Updates", "5360fd7a-a1c4-428f-91c9-89a4c5565ce1"),
    ("All_Windows_10_Updates", "6d833cef-aa77-4018-943c-627979f5905c") | ForEach-Object { [pscustomobject]@{Name = $_[0]; Guid = $_[1] } }




    Foreach ($Guid in $Guids) {
        Write-Host "Create Extension for $($Guid.Name)"
        $GuidFolder = New-Item -Path "$($ENV:SMS_ADMIN_UI_PATH)\..\..\XmlStorage\Extensions\Actions" -Name "$($Guid.Guid)" -ItemType "directory" -Force
        $TempXML = $null
        [xml]$TempXML = $Extension
        $TempXML.SelectNodes("//Parameters") | % { 
            $_."#text" = $_."#text".Replace("#PATHTOSCRIPT#", $PathToScript)
            $_."#text" = $_."#text".Replace("#SQLSERVER#", $SQLServer)
            $_."#text" = $_."#text".Replace("#SUSDB#", $SQLDBName)
            if ($($Guid.Name) -eq "All_Windows_10_Updates") {
                $PARAMS = '-UpdateSearchString "##SUB:LocalizedDisplayName##"'
            }
            else {
                $PARAMS = '-UpdateID "##SUB:CI_UniqueID##" -RevisionNumber "##SUB:RevisionNumber##"'
            }
            $_."#text" = $_."#text".Replace("#PARAMS#", $PARAMS)
        }    
        
        $TempXML.Save("$GuidFolder\ShowApplicabilityRule.xml")

    }
}