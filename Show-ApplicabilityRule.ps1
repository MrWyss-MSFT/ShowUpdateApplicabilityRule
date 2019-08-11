<#
.SYNOPSIS
    Shows Applicability Rule for a given Update
.DESCRIPTION
    Return the Applicability Rule of a given Update. The Applicability Rules are stored in the
    WSUS Database (compressed blob datatype image) as an XML File in the field RootElementXmlCompressed
    in tbxml.
    This script will search the an update by either a given SearchString or UpdateID and RevisionNumber.
    Once found, it will find the updates xml. If the found update is a boundle it will use a recurse function
    to find the bundle updates XML. Once found the XML needs to be extracted (cab) and parsed to get to the 
    Applicability Rule.
    
.NOTES  
    File Name   : Show-ApplicabilityRule.ps1
    Author      : marius.wyss@microsoft.com
    Version     : 3.0
    ChangeLog   : 29. Jul 2019 - V3.0 - Introduced Parameter Set to avoid empty UpdateSarchstring or emtpy UpdateID and RevisionNumber
                  27. May 2019 - V2.0 - Added TreeView thanks to Ivan Yankulov https://www.sptrenches.com/2017/08/build-treevie-xml-powershell.html
                  27. May 2019 - V1.1 - minor refactoring and invoke-sqlcmd2 got error handling to avoid open connections
                  26. May 2019 - V1.0 - initial version

    THIS SAMPLE CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
    WHETHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
    WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
    IF THIS CODE AND INFORMATION IS MODIFIED, THE ENTIRE RISK OF USE OR RESULTS IN
    CONNECTION WITH THE USE OF THIS CODE AND INFORMATION REMAINS WITH THE USER.
#>
[CmdletBinding(DefaultParametersetName='Script')] 
param( 
    # Use UpdateSearchString or UpdateID and RevisionNumber to find the Update
    [Parameter(Position=0,Mandatory=$true)] [string]$SQLServer, # e.g. = sqlserver.domain.com
    [Parameter(Position=1,Mandatory=$true)] [string]$SQLDBName, # e.g. = SUSDB
    [Parameter(ParameterSetName='Script',Mandatory=$true)][string]$UpdateSearchString, # e.g. = "%Office 365 Client Update - First Release for Current Channel Version 1706 for x64 based Edition (Build 8229.2056)%",     
    [Parameter(ParameterSetName='Extension',Mandatory=$true)][string]$UpdateID,
    [Parameter(ParameterSetName='Extension',Mandatory=$true)][string]$RevisionNumber
)

$ParamSetName = $PsCmdLet.ParameterSetName
Function Invoke-Sqlcmd2 {
    [CmdLetBinding()]
    Param
    (
        [Parameter(Mandatory = $True)]
        [String]$Server,
        [Parameter(Mandatory = $True)]
        [String]$Database,
        [Parameter(Mandatory = $True)]
        [String]$SQLQuery

    )
    Try {
        # Prepare SQL-Connection-Object
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.StatisticsEnabled = $False
        $SqlConnection.ConnectionString = "Server = $Server; Database = $Database; Integrated Security = True"
        $SqlConnection.Open()
    
        # Prepare SQL-Command-Object
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SqlQuery
        $SqlCmd.Connection = $SqlConnection

 
        # Prepare SQL-Adapter-Object for the Select-Exectuion
        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd
 
        # Prepare DataSet with the values of the Select-Statement
        $Dataset = New-object System.Data.Dataset
        $SqlAdapter.Fill($Dataset)
        $Dataset.Tables[0]    
    }
    Catch {
        Write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        Write-host "Exception Stack: $($_.ScriptStackTrace)"
    }
    Finally {
        #Write-Host "Get Stats, Close Connection and Release all Connection Ressources"
        #$SqlConnection.RetrieveStatistics() | Format-Table
        $SqlConnection.Close()
        $SqlConnection.Dispose()
        $SQLQuery = $null
    }
}

Function Get-UpdateIDAndRevision {
    Param ([string]$UpdateSearchString)
    $sql_GetUpdateAndRevision = @'
DECLARE @SearchString AS Nvarchar(255) = '{0}';
DECLARE @UpdateID AS uniqueidentifier;
DECLARE @RevisionNumber AS int;
SELECT * FROM PUBLIC_VIEWS.vUpdate WHERE DefaultTitle like @SearchString and IsDeclined = 0
'@ -f $UpdateSearchString
          
    $UpdAndRev = Invoke-Sqlcmd2 -Server $SQLServer -Database $SQLDBName -SQLQuery $sql_GetUpdateAndRevision
    $UpdAndRev

}
Function Get-UpdateDetails {
    Param ([string]$UpdateID, [String]$RevisionNumber)
    $sql_GetUpdateAndRevision = "SELECT * FROM PUBLIC_VIEWS.vUpdate WHERE UpdateID='{0}' and RevisionNumber = '{1}'" -f $UpdateID, $RevisionNumber
    $Update = Invoke-Sqlcmd2 -Server $SQLServer -Database $SQLDBName -SQLQuery $sql_GetUpdateAndRevision
    $Update

}

#$level = 0
function Get-UpdateXML {
    param(  [string]$UpdateID,
        [string]$UpdateRevision
    )
    $level++
    if ($level -ge 2) { write-host "Bundle Update" }
    
    $sql_getxml = @'
DECLARE @XMLID AS int;
--DECLARE @FILENAME as nvarchar(15) = 'Temp.cab';
SELECT @XMLID = XmlID  FROM dbo.tbXml x 
INNER JOIN dbo.tbRevision r ON x.RevisionID = r.RevisionID
INNER JOIN dbo.tbUpdate u ON r.LocalUpdateID = u.LocalUpdateID
WHERE u.UpdateID = '{0}' AND r.RevisionNumber = {1} And RootElementType = '0'
--SELECT @FILENAME as filename, ISNULL(datalength(RootElementXmlCompressed), 0) size, RootElementXmlCompressed FROM SUSDB.dbo.tbxml WHERE XmlID = @XMLID
SELECT ISNULL(datalength(RootElementXmlCompressed), 0) size, RootElementXmlCompressed FROM SUSDB.dbo.tbxml WHERE XmlID = @XMLID
'@ -f $UpdateID, $UpdateRevision
    $UpdateXML = Invoke-Sqlcmd2 -Server $SQLServer -Database $SQLDBName -SQLQuery $sql_getxml

    #Dirty TempFile Trick
    #region TODO: Tempfile Workaround
    
    $tmpfile = New-TemporaryFile
    [System.Text.Encoding]::default.GetString($UpdateXML.RootElementXmlCompressed) | Out-File $tmpfile -Encoding default 

  
    if (C:\Windows\System32\expand.exe) {
        $xml = "$($tmpfile.DirectoryName)\blob.xml"
        try { cmd.exe /c "C:\Windows\System32\expand.exe -F:* $($tmpfile.FullName) $($xml)" | out-null }
        catch { Write-host "error expanding" }
    }

    Remove-Item $($tmpfile.FullName) -Force

    [System.Xml.XmlDocument]$doc = new-object System.Xml.XmlDocument;
    $doc.set_PreserveWhiteSpace( $true );
    $doc.Load( $xml );
    Remove-Item $xml -Force

    $root = $doc.get_DocumentElement();
    $xml = $root.get_outerXml();
    $xml = '<?xml version="1.0" encoding="utf-8"?>' + $xml

    $newFile = $($tmpfile.FullName) + ".utf8.xml"
    Set-Content -Encoding UTF8 $newFile $xml;
    $utf8xml = [xml](Get-Content $newFile)
    # Refactoring
    If ($null -ne $utf8xml.Update.Relationships.BundledUpdates.UpdateIdentity) {
        Write-Debug "Bundled Update"
        [string]$UpdateID = $utf8xml.Update.Relationships.BundledUpdates.UpdateIdentity.UpdateId
        [string]$UpdateRevisionNumber = $utf8xml.Update.Relationships.BundledUpdates.UpdateIdentity.RevisionNumber
                 
        $utf8xml = Get-UpdateXML -UpdateID $UpdateID -UpdateRevision $UpdateRevisionNumber
    
    }
    Remove-Item $newFile -Force
    #endregion
    return [System.Xml.XmlDocument]$utf8xml
}
function Format-XML ([xml]$xml) {
    $StringWriter = New-Object System.IO.StringWriter;
    $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter;
    $XmlWriter.Formatting = "indented";
    $xml.WriteTo($XmlWriter);
    $XmlWriter.Flush();
    $StringWriter.Flush();
    Write-Output $StringWriter.ToString();
}
Function Get-ApplicabilityRules {
    param ([System.Xml.XmlDocument]$UpdateXML)
    #$isinstalled = $UpdateXML.Update.ApplicabilityRules.IsInstalled.InnerXml
    #$isInstallable = $UpdateXML.Update.ApplicabilityRules.IsInstallable.InnerXml

    $xmlstring += $UpdateXML.Update.ApplicabilityRules.InnerXml
    
    #[xml]$ar = "<ApplicabilityRules><IsInstalled>$($isinstalled)</IsInstalled><IsInstallable>$($isInstallable)</IsInstallable></ApplicabilityRules>"
    [xml]$ar = "<ApplicabilityRules>$($xmlstring)</ApplicabilityRules>"
    Format-XML -xml $ar 

}

function Switch-View () {
    if ($Button1.text -eq "Switch to TextView") {
        $tbARText.Width = $TREEVIEW.Width
        $tbARText.Height = $TREEVIEW.Height
        $Form.controls.Remove($TREEVIEW)
        $Form.controls.Add($tbARText)
        $Button1.text = "Switch to TreeView"
        $Label1.text = $TextViewText
    }
    else {
        $TREEVIEW.Width = $tbARText.Width
        $TREEVIEW.Height = $tbARText.Height
        $Form.controls.Remove($tbARText)
        $Form.controls.Add($TREEVIEW)
        $Button1.text = "Switch to TextView"
        $Label1.text = $TreeViewText
    }
}


function Add-NodesToTreeview($xElement, $fNode) {
    $allEl = $xElement.ChildNodes
    foreach ($xEl in $allEl) {
        $tn = new-object System.Windows.Forms.TreeNode
        if ($xEl.HasChildNodes) {
            $tn.Text = $xEl.OuterXml.Split(">")[0] + ">"
            $tn.Tag = $xEl.OuterXml
            [void]$fNode.Nodes.Add($tn)
            Add-NodesToTreeview -xElement $xEl -fNode $tn
        }
        else {
            $tn.Text = $xEl.OuterXml
            $tn.Tag = $xEl.OuterXml
            [void]$fNode.Nodes.Add($tn)
        }     
    }


}


#region Business Logic
Try {
    
    if ($ParamSetName -eq "Extension") {
        Write-Host "Using UpdateID and Revision"
        $Update = Get-UpdateDetails -UpdateID $UpdateID -RevisionNumber $RevisionNumber
     
    }
    elseif ($ParamSetName -eq "Script") {
        Write-Host "Using SearchString"
        $Update = Get-UpdateIDAndRevision -UpdateSearchString $UpdateSearchString
        [string]$UpdateID = $Update.UpdateID
        [string]$RevisionNumber = $Update.RevisionNumber
    } else {
        Write-Host "No Parameters given"
    }

    
    $UpdateXML = Get-UpdateXML -UpdateID $UpdateID -UpdateRevision $RevisionNumber
    $Ar = Get-ApplicabilityRules -Update $UpdateXML

    #region ShowDialog  
    $RootElementName = "ApplicabilityRules"
    $xml = [xml]$Ar
    
    
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $FormWidth = 1024
    $FormHeight = 730
    $SidePanel = 300
    $TreeViewText = "Treeview: Double Click with copy entire node."
    $TextViewText = "TextView: select and copy."

    $FORM = new-object Windows.Forms.Form    
    $FORM.text = "Form-DisplayDirExplorer"
    $FORM = new-object Windows.Forms.Form
    $FORM.Size = new-object System.Drawing.Size($FormWidth, $FormHeight)    
    $FORM.text = "Update XML"


    $TREEVIEW = new-object windows.forms.TreeView 
    $TREEVIEW.Width = $($FormWidth - $SidePanel)
    $TREEVIEW.Height = $FormHeight
    $TREEVIEW.Anchor = "right, left, top, bottom"
    $TREEVIEW.AutoSize = $false

    $tbARText = New-Object system.Windows.Forms.TextBox
    $tbARText.multiline = $true
    $tbARText.AutoSize = $false
    $tbARText.text = $xml.InnerXml
    $tbARText.width = $($FormWidth - $SidePanel)
    $tbARText.height = $FormHeight
    $tbARText.Anchor = "right, left, top, bottom"
    $tbARText.Font = 'Microsoft Sans Serif,8'

    $Label1 = New-Object system.Windows.Forms.Label
    $Label1.text = $TreeViewText
    $Label1.AutoSize = $true
    $Label1.width = $($SidePanel - 10)
    $Label1.height = 10
    $Label1.Anchor = "right, top"
    $Label1.location = New-Object System.Drawing.Point($(($FormWidth - $SidePanel + 10)), 10)
    $Label1.Font = 'Microsoft Sans Serif,10'

    $Button1 = New-Object system.Windows.Forms.Button
    $Button1.text = "Switch to TextView"
    $Button1.width = 180
    $Button1.height = 30
    $Button1.Anchor = "right, top"
    $Button1.location = New-Object System.Drawing.Point($(($FormWidth - $SidePanel + 10)), 100)
    $Button1.Font = 'Microsoft Sans Serif,10'
    $Button1.Add_Click( { Switch-View $this $_Test })

    $TREEVIEW.add_NodeMouseDoubleClick( {
            [Windows.Forms.Clipboard]::SetText($TREEVIEW.SelectedNode.Tag.ToString());
        })

    $firstElement = $xml.$RootElementName
    $tn = new-object System.Windows.Forms.TreeNode
    $tn.Text = $firstElement.OuterXml.Split(">")[0] + ">"
    $tn.Tag = $firstElement.OuterXml

    [void]$TREEVIEW.Nodes.Add($tn)

    Add-NodesToTreeview -xElement $firstElement -fNode $tn
    $Form.controls.AddRange(@($TREEVIEW, $Label1, $Button1))

    $Form.ShowDialog()
}
catch {
    Write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-host "Exception Stack: $($_.ScriptStackTrace)"
    Write-Host -NoNewLine 'Press any key to continue...';
    #$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}