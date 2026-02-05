###########################################################################################
########################### Number of HealthMailboxes #####################################
############################# and corrective actions  #####################################
###########################################################################################
#
# Copyright - Steffen Meyer
#
# Change History
# 03/09/2024 V1.0 Initial version
# 04/09/2024 V1.1 elseif statement missing + Error/WarningAction handling
# 02/05/2026 V1.2 changed duplicate mailboxes handling
#
###########################################################################################

$version = "1.2"

write-host "----------------------------------------------------------------------------"
write-host "|              This script will compare the number                         |"
write-host "|         of given and expected HealthMailboxes in your                    |" 
write-host "|         Exchange organization and if needed will try                     |"
write-host "|                  to correct a wrong number.                              |"
write-host "----------------------------------------------------------------------------"

write-host "`nScriptversion: $version"

#Check if Exchange SnapIn is available and load it
if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange")
{
    if ((Get-PSSnapin -Registered).name -contains "Microsoft.Exchange.Management.PowerShell.SnapIn")
    {
        Write-Host "`nLoading the Exchange Powershell SnapIn..." -ForegroundColor Yellow
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
    Write-Host "`nExchange Management Tools are not installed. Run the script on a different machine." -ForegroundColor Red
    Return
    }
}

#Detect, where the script is executed
if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue))
{
    write-host "`nATTENTION: Script is executed on a non-Exchangeserver...`n" -ForegroundColor Cyan
}

function Get-HealthMailboxes
{
    #Number of HealthMailboxes in org
    Set-ADServerSettings -viewentireforest $true
    try
    {
        Get-Mailbox -Monitoring -Resultsize unlimited -IgnoreDefaultScope -WarningAction SilentlyContinue -ErrorAction Stop
    }
    catch
    {
        Write-Host "`nThere were some issues while trying to find HealthMailboxes in this organization." -ForegroundColor Red
        return
   }
}

#Count expected HealthMailboxes for ClientAccess
Write-Host "`nCounting ClientAccess roles..." -ForegroundColor White
[int]$casmbx = (Get-ClientAccessService).count * 10

#Count expected HealthMailboxes for DBCopies (no recovery mdbs)
Write-Host "`nCounting MailboxDatabaseCopies, this may take a while..." -ForegroundColor White
[int]$dbcopymbx = (Get-MailboxDatabase | Where-Object Recovery -ne $true | Get-MailboxDatabaseCopyStatus | Measure-Object).count

#TotalExpected
[int]$expmbx = $casmbx + $dbcopymbx

Write-Host "`nCounting HealthMailboxes, this may take a while..." -ForegroundColor White
[array]$mailboxes = Get-HealthMailboxes
[int]$tot = ($mailboxes).count

Write-Host "`nSUM:" -Foregroundcolor Yellow
Write-Host "We found $($tot) HealthMailboxes in this organization." -Foregroundcolor White

if ($expmbx -eq $tot)
{
    Write-Host "`nRESULT:" -Foregroundcolor Green
    Write-Host "The number of $($tot) HealthMailboxes is correct and matches the expected number of $($expmbx), no further actions needed.`n" -ForegroundColor Green
    return
}
elseif ($expmbx -gt $tot)
{
    $diff = $expmbx - $tot
    Write-Host "`nRESULT:" -Foregroundcolor Yellow
    Write-Host "The number of HealthMailboxes of $($tot) is lower than the expected number of $($expmbx), Exchange Health Manager service needs to create the missing ones." -ForegroundColor Yellow
    #$n = 1
    return
}
else
{
    $diff = $tot - $expmbx
    Write-Host "`nRESULT:" -Foregroundcolor Yellow
    Write-Host "The number of HealthMailboxes of $($tot) is greater than the expected number of $($expmbx), we assume there are $($diff) duplicate or outdated HealthMailboxes." -ForegroundColor Yellow
    $n = 2
}

#Corrective actions, when HealthMailboxes are missing
if ($n -eq 1)
{
    #usually, Exchange healthmanager service will add missing mailboxes automatically
    #corrective actions still needs to be defined, waiting for an example
}

#Corrective actions, when there are too many HealthMailboxes in the organization
if ($n -eq 2)
{
    #First detecting and deleting duplicates
    $duplicates = $mailboxes | Group-Object displayname | Where-Object count -gt 1
        
    if ($duplicates)
    {
        $dupsum = ($duplicates| Measure-Object).count
        $duptodelete = ($duplicates | Measure-Object -sum count).sum - $dupsum
                
        $removedup = Read-Host "`nNOTICE: We found $($duptodelete) duplicate mailboxes, continue to delete duplicate HealthMailboxes? ( Y / N )"
        
        if  ($removedup -ne "Y")
        {
            Write-Host "NOTICE: No duplicate HealthMailboxes were deleted.`n" -ForegroundColor Yellow
        }
        else
        {
            Write-Host "`nNOTICE:`n$($dupsum) mailbox names with duplicates will be processed and $($duptodelete) duplicate HealthMailboxes will be removed now!" -ForegroundColor Yellow
    
            #Counter for progress bar
            $i = 0

            ForEach ($dupgroup in $duplicates)
            {
                #Progress bar
                $i++
                $status = "{0:N0}" -f ($i / $dupsum * 100)
                Write-Progress -Activity "Removing duplicate HealthMailboxes..." -Status "Processing duplicate mailboxname $i of $dupsum : $status% Completed" -PercentComplete ($i / $dupsum * 100)
                                
                $dupmbxs = $dupgroup.group | foreach { Get-Mailbox -Monitoring $_.exchangeguid.guid -WarningAction SilentlyContinue -ErrorAction SilentlyContinue }
                $dupmbxs = $dupmbxs | Sort-Object whencreated -Descending | Select-Object -Skip 1

                foreach ($dupmbx in $dupmbxs)
                {
                    $dupremove = Remove-Mailbox $dupmbx.exchangeguid.guid -confirm:$false
                }
            }
        Write-Progress -Completed -Activity "Done!"
        Write-Host "`nCleanup of duplicate HealthMailboxes done!" -ForegroundColor Green
        }
    }
        
    #Second detecting and deleting old HealthMailboxes for dismantled resources
    $olds = @()        
    
    #Exchange Servers
    [array]$servers = (Get-Exchangeserver | Sort-Object fqdn).name
    
    #HealthMailbox displayname (may consist of an old/dismantled servername)
    [array]$mbxnames = ($mailboxes).displayname
            
    foreach ($mbxname in $mbxnames)
    {
        if ($mbxname -notmatch ($servers -join '|' ))
        {
            $olds += $mbxname
        }
    }
    
    $oldsum = $olds.count

    if ($olds)
    {
        $removeold = Read-Host "`nNOTICE:We found $($oldsum) old/outdated HealthMailboxes (some may overlap with duplicates), delete it? ( Y / N )"

        if ($removeold -ne "Y")
        {
            Write-Host "NOTICE: No old/outdated HealthMailboxes were deleted.`n" -ForegroundColor Yellow
            return
        }
        else
        {
            Write-Host "`nNOTICE:`n$($oldsum) old/outdated HealthMailboxes will be removed now!" -ForegroundColor Yellow
           
            $i = 0

            ForEach ($old in $olds)
            {
                #Progress bar
                $i++
                $status = "{0:N0}" -f ($i / $oldsum * 100)
                Write-Progress -Activity "Removing old/outdated HealthMailboxes..." -Status "Processing mailbox $i of $oldsum : $status% Completed" -PercentComplete ($i / $oldsum * 100)

                #Old/outdated mailbox removal
                Get-Mailbox -Monitoring $old.exchangeguid.guid -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Remove-Mailbox -confirm:$false
            }
            Write-Progress -Completed -Activity "Done!"
            Write-Host "`nCleanup of old/outdated HealthMailboxes done!" -ForegroundColor Green
        }
    }
    elseif ($removedup -ne "Y" -and !($olds))
    {
    return
    }
}

#Waiting 30 seconds for AD replication
Write-Host "`nNOTICE: Waiting 30 seconds for AD replication..." -ForegroundColor Yellow
Start-Sleep 30

#Count again
Write-Host "`nCounting HealthMailboxes again, this may take a while..." -ForegroundColor White
[int]$tot = (Get-HealthMailboxes).count
    
if ($expmbx -eq $tot)
{
    Write-Host "`nAfter the cleanup, the number of $($tot) HealthMailboxes now matches the expected number of $($expmbx), no further actions needed.`n" -ForegroundColor Green
}
else
{
    Write-Host "`nAfter the cleanup, the number of $($tot) HealthMailboxes still does not match the expected number of $($expmbx), you need to re-run the script and/or take further actions.`n" -ForegroundColor Red
}
#END
