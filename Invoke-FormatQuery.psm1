<#
 .Synopsis
  Return PsCustumObject form a query in sccm

 .Description
  Run & Format a SCCM query to a powershell object usable whit all command & module (like Out-GridView)

 .Parameter SiteCode
  The Site Code of SCCM Site

 .Parameter ProviderMachineName
  The ProviderMachineName of SCCM Site

  .Parameter QueryId
  The ID of SCCM query

 .Parameter LimitId
  Id of collection to limit the query

  .Parameter wql
  The wql to run in the query

 .Example
  Invoke-FormatQuery "SI1" "provider.test.exemple.com" "CO100567" "CO100342"

#>

function Invoke-FormatQuery()
{
    param([Parameter(Mandatory=$true)]
        [string]$SiteCode,
        [Parameter(Mandatory=$true)]
        [string]$ProviderMachineName,
        [Parameter(Mandatory=$true)]
        [string]$QueryId,
        [Parameter(Mandatory=$false)]
        [string]$LimitId,
        [Parameter(Mandatory=$false)]
        [switch]$formatOU = $false
    )
    $initParams = @{}
    if($null -eq (Get-Module ConfigurationManager)){Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams}
    if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams}
    Set-Location "$($SiteCode):\" @initParams

    if($LimitId.Length -gt 5)
    {
        $object = Invoke-CMQuery -Id $QueryId -LimitToCollectionId $LimitId
    }
    else
    {
        $object = Invoke-CMQuery -Id $QueryId
    }
    $size = $object.count
    $table = @()
    $table_temp = ""

    for($i = 0;$i -le $size;$i++)
    {
        $per = [math]::Round(($i*100)/$size)
        Write-Progress -Activity "Analysing SCCM Data" -PercentComplete $per -Status "$per`%"
        [string[]]$root_list = $object[$i].PropertyList.Keys | ForEach-Object{$_}
        foreach($property in $root_list)
        {
            [string[]]$work_property = $object[$i].$property.PropertyList.Keys | ForEach-Object{$_}

            foreach($line in $work_property)
            {
                [string[]]$value = $object[$i].$property.$line
                if($formatOU -eq $true)
                {
                    if($line -like "*OuName*")
                    {
                        if($value.Count -gt 2)
                        {
                            $mc = ($value.Count)-1
                            $value = $value[$mc]
                        }
                    }
                }
                
                if($value -like "*+000")
                {
                        $temp = "";
                        [string]$temp = $value
                
                        $temp = $temp.Substring(0,$temp.Length-11)
                        $temp = $temp.Insert(4,"-")
                        $temp = $temp.Insert(7,"-")
                        $temp = $temp.Insert(10," ")
                        $temp = $temp.Insert(13,"h")
                        $temp = $temp.Insert(16,"m")
                        $temp = $temp.Insert(19,"s")
                        $value = $temp
                        
                    }
                $myadd = "$line`=$value`n"
                $myline += "$myadd"
            }
        }

$table_temp = @"
$myline
"@
$myline = ""
    $table_hash = ConvertFrom-StringData -StringData $table_temp 
    $table += [Pscustomobject]$table_hash
    }
    return $table
}

function Invoke-FormatWQLQuery()
{
    param([Parameter(Mandatory=$true)]
        [string]$SiteCode,
        [Parameter(Mandatory=$true)]
        [string]$ProviderMachineName,
        [Parameter(Mandatory=$true)]
        [string]$wql
    )
    $initParams = @{}
    if($null -eq (Get-Module ConfigurationManager)){Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams}
    if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams}
    Set-Location "$($SiteCode):\" @initParams

    $object = Invoke-CMWmiQuery -Query $wql
    $size = $object.count
    $table = @()
    $table_temp = ""

    for($i = 0;$i -le $size;$i++)
    {
        $per = [math]::Round(($i*100)/$size)
        Write-Progress -Activity "Analysing SCCM Data" -PercentComplete $per -Status "$per`%"
        [string[]]$root_list = $object[$i].PropertyList.Keys | ForEach-Object{$_}
        foreach($property in $root_list)
        {
            [string[]]$work_property = $object[$i].$property.PropertyList.Keys | ForEach-Object{$_}

            foreach($line in $work_property)
            {
                [string[]]$value = $object[$i].$property.$line
                if($line -like "*OuName*")
                {
                    if($value.Count -gt 2)
                    {
                        $mc = ($value.Count)-1
                        $value = $value[$mc]
                    }
                }
                if($value -like "*+000")
                {
                        $temp = "";
                        [string]$temp = $value
                
                        $temp = $temp.Substring(0,$temp.Length-11)
                        $temp = $temp.Insert(4,"-")
                        $temp = $temp.Insert(7,"-")
                        $temp = $temp.Insert(10," ")
                        $temp = $temp.Insert(13,"h")
                        $temp = $temp.Insert(16,"m")
                        $temp = $temp.Insert(19,"s")
                        $value = $temp
                        
                    }
                $myadd = "$line`=$value`n"
                $myline += "$myadd"
            }
        }

$table_temp = @"
$myline
"@
$myline = ""
    $table_hash = ConvertFrom-StringData -StringData $table_temp 
    $table += [Pscustomobject]$table_hash
    }
    return $table
}

# Export functions
Export-ModuleMember -Function @('Invoke-FormatWQLQuery','Invoke-FormatQuery')