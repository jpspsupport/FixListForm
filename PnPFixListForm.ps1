<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 

 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>


param(
  [Parameter(Mandatory=$true)]
  $ApplicationId,
  [Parameter(Mandatory=$true)]
  $siteUrl,
  [Parameter(Mandatory=$true)]
  $listName,
  [ValidateSet("ALL", "DISPLAY", "EDIT", "NEW")]
  $formtype = "ALL",
  [switch]$force
)
$ErrorActionPreference = "Stop"

Connect-PnPOnline $siteUrl -ApplicationId $ApplicationId -Interactive
# This might not work, but calling just in case
Set-PnPSite -NoScriptSite $false
$context = Get-PnPContext

$web = Get-PnPWeb
$list = $web.Lists.GetByTitle($listName)
$context.Load($list)
$context.ExecuteQuery()
$parentFolder = $list.RootFolder
$context.Load($parentFolder)
$context.ExecuteQuery()
if ($list.BaseType -eq [Microsoft.SharePoint.Client.BaseType]::DocumentLibrary)
{
  $formsurl = $parentFolder.ServerRelativeUrl + "/Forms"
  $parentFolder = $web.GetFolderByServerRelativeUrl($formsurl)
  $context.Load($parentFolder)
  $context.ExecuteQuery()
}
function CreateView($fileName, $ControlMode, $FormType)
{
  $url = $parentFolder.ServerRelativeUrl + "/" + $fileName
  $fileType = [Microsoft.SharePoint.Client.TemplateFileType]::FormPage
  $checkfile = $web.GetFileByServerRelativeUrl($url)
  $context.Load($checkfile)
  try
  {
    $context.ExecuteQuery()
    $fileExists = $checkfile.Exists
  }
  catch{}
  if ($force)
  {
    if ($fileExists)
    {
      $checkfile.DeleteObject()
      $context.ExecuteQuery()
      $fileExists = $false
      Write-Host "The page" $url "is removed."      
    }
  }
  if (!$fileExists)
  {
    $file = $parentFolder.Files.AddTemplateFile($url, $fileType)
    $context.Load($file)
    $context.ExecuteQuery()
    Write-Host "The page" $url "is created."      
  }

  $wpm = $file.GetLimitedWebPartManager([Microsoft.SharePoint.Client.WebParts.PersonalizationScope]::Shared)
  $webPartXml = '<?xml version="1.0" encoding="utf-8"?>'
  $webPartXml += '<WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">'
  $webPartXml += '<Title>' + $list.Title + '</Title>'
  $webPartXml += '<FrameType>Default</FrameType>'
  $webPartXml += '<Description/>'
  $webPartXml += '<IsIncluded>true</IsIncluded>'
  $webPartXml += '<PartOrder>2</PartOrder>'
  $webPartXml += '<FrameState>Normal</FrameState>'
  $webPartXml += '<Height/>'
  $webPartXml += '<Width/>'
  $webPartXml += '<AllowRemove>true</AllowRemove>'
  $webPartXml += '<AllowZoneChange>true</AllowZoneChange>'
  $webPartXml += '<AllowMinimize>true</AllowMinimize>'
  $webPartXml += '<AllowConnect>true</AllowConnect>'
  $webPartXml += '<AllowEdit>true</AllowEdit>'
  $webPartXml += '<AllowHide>true</AllowHide>'
  $webPartXml += '<IsVisible>true</IsVisible>'
  $webPartXml += '<DetailLink/>'
  $webPartXml += '<HelpLink/>'
  $webPartXml += '<HelpMode>Modeless</HelpMode>'
  $webPartXml += '<Dir>Default</Dir>'
  $webPartXml += '<PartImageSmall />'
  $webPartXml += '<MissingAssembly>Cannot import this Web Part.</MissingAssembly>'
  $webPartXml += '<PartImageLarge/>'
  $webPartXml += '<IsIncludedFilter/>'
  $webPartXml += '<Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>'
  $webPartXml += '<TypeName>Microsoft.SharePoint.WebPartPages.ListFormWebPart</TypeName>'
  $webPartXml += '<ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{' + $list.ID.ToString() + '}</ListName>'
  $webPartXml += '<ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">' + $list.ID.ToString() + '</ListId>'
  $webPartXml += '<ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">' + $ControlMode + '</ControlMode>'
  $webPartXml += '<TemplateName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">ListForm</TemplateName>'
  $webPartXml += '<FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">' + $FormType + '</FormType>'
  $webPartXml += '<ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">1048576</ViewFlag>'
  $webPartXml += '</WebPart>'
  $webPartDefinition = $wpm.ImportWebPart($webPartXml)
  $webPart = $webPartDefinition.WebPart
  [void]$wpm.AddWebPart($webPart, "Main", 1)
  $context.Load($webPart)
  $context.ExecuteQuery()
  Write-Host "The webpart on " $fileName " is added"
}
$formtype = $formtype.ToUpper()
if (($formtype -eq "ALL") -or ($formtype -eq "DISPLAY"))
{
  CreateView -fileName "DispForm.aspx" -ControlMode "Display" -FormType "4"
}
if (($formtype -eq "ALL") -or ($formtype -eq "EDIT"))
{
  CreateView -fileName "EditForm.aspx" -ControlMode "Edit" -FormType "6"
}
if (($formtype -eq "ALL") -or ($formtype -eq "NEW"))
{
  CreateView -fileName "NewForm.aspx" -ControlMode "New" -FormType "8"
}