$O_Draft_Folder = 16
$O_Object = New-Object -comObject Outlook.Application
$O_Namespace = $O_Object.GetNameSpace("MAPI")
$Get_Drafts = $O_Namespace.GetDefaultFolder($O_Draft_Folder)
Function New-Draft{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Subject,
        [string]$Body,
        [string]$Recipients
    )
    $New_Draft = $O_Object.CreateItem(0)
    $New_Draft.Subject = $Subject
    $New_Draft.Recipients.Add($Recipients) | Out-Null
    $New_Draft.Body = $Body
    $New_Draft.save()
}
Function Set-Draft{
    Param(
        [Parameter(Mandatory)]
        [string]$bySubject,
        [string]$AddSubject,
        [string]$ChangeSubject,
        [string]$AddBody,
        [string]$ChangeBody,
        [string]$AddRecipients
    )
    try{
        $Get_Draft = $Get_Drafts.Items | where {$_.Subject -eq $bySubject}
    }catch{
        Write-Output "Item could not be found"
    }   
    ##
    if($AddSubject -and $ChangeSubject){
        Write-Error "You cannot Add and Change at the same time"
    }elseif($AddSubject){
        $Get_Draft.Subject += $AddSubject
    }elseif($ChangeSubject){
        $Get_Draft.Subject = $ChangeSubject
    }
    ###
    if($AddBody -and $ChangeBody){
        Write-Error "You cannot Add and Change at the same time"
    }elseif($AddBody){
        $Get_Draft.Body += $AddBody
    }elseif($ChangeBody){
        $Get_Draft.Body = $ChangeBody
    }
    ###
    if($AddRecipients){
        $Get_Draft.Recipients.Add($AddRecipients) | Out-Null
    }
    $Get_Draft.save()
}
Function Get-Draft{
    Param(
        [Parameter(Mandatory)]
        [string]$Subject
    )
    try{
        $Get_Draft = $Get_Drafts.Items | where {$_.Subject -eq $Subject}
    }catch{
        Write-Output "Item could not be found"
    }
    return $Get_Draft
}
Function Send-Draft{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$bySubject
    )
    try{
        $Get_Draft = $Get_Drafts.Items | where {$_.Subject -eq $bySubject}
    }catch{
        Write-Output "Item could not be found"
    }
    $Get_Draft.Send()
}
#to use this script it needs to be ran with Windows Powershell of the same version (32bit/64bit) as the version of your Outlook