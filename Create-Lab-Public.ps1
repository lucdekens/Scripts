#region Start
$credentials = "$($env:USERPROFILE)\ravello-credentials.csv"

$obj = Import-Csv -Path $credentials -UseCulture
$sPswd = ConvertTo-SecureString -String $obj.Pswd -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($obj.User, $sPswd)

$sConn = @{
  Credential = $cred
}
$connect = Connect-Ravello @sConn
#endregion

#region Payload
$lab = Get-Content yggdrasil.json | Out-String | ConvertFrom-Json

Write-Host "`r`r"
Get-Date -Format 'hh:mm:ss'

# Upload ISOs
$lab.ISO | %{
  if(!(Get-RavelloDiskImage -DiskImageName $_.Filename))
  {
    Import-Ravello -IsoPath "$($_.Path,$_.Filename -join '\')" -Confirm:$false
    $doc = "**$($_.Label)**: uploaded $(Get-Date -Format 'yyyyMMdd-hhmm') by $($env:USERNAME) from $($env:COMPUTERNAME)"
    Get-RavelloDiskImage -DiskImageName $_.Filename | New-RavelloDiskImageDocumentation -Documentation $doc -Confirm:$false
  }
}

# Create the Application
$lab.Lab | %{
  $app = Get-RavelloApplication -ApplicationName $_.LabName

  # Remove previous instance when Force
  if($app -and [Convert]::ToBoolean($_.Force))
  {
    Remove-RavelloApplication -ApplicationId $app.id -Confirm:$false
    $app = $null
    Write-Host 'Removed previous instance of application ' -NoNewline
    Write-Host -ForegroundColor Green "$($_.LabName)"
  }
  
  if(!$app)
  {
    Write-Host 'Creating Application ' -NoNewline
    Write-Host -ForegroundColor Green "$($_.LabName)"
    
    $order = @()
    $sApp = @{
      ApplicationName = $_.LabName
      Description = $_.LabDescription
      Confirm = $false
    }
    $doc = "**$($_.Labname)**: created on $(Get-Date -Format 'yyyyMMdd-hhmm') by $($env:USERNAME) from $($env:COMPUTERNAME)`r"

    # Create Application and add Documentation
    $app = New-RavelloApplication @sApp
    $app | New-RavelloApplicationDocumentation -Documentation $doc -Confirm:$false | Out-Null
      
    # Add the VMs
    $lab.VM | where{$_.LabName -eq $app.name} | %{
      Write-Host "`tAdding VM $($_.VmName)..." -NoNewline
      $sAddVm = @{
        ApplicationId = $app.id
        ImageName = $_.Image
        NewVmName = $_.VmName
        NewVmDescription = $_.VmDescription
        Confirm = $false
      }
      $newLine = "- **$($_.VmName)**: added to $($app.name) from $($_.Image)" +
      " on $(Get-Date -Format 'yyyyMMdd-hhmm') by $($env:USERNAME) from $($env:COMPUTERNAME)"
      $doc = (Get-RavelloApplicationDocumentation -ApplicationId $app.id) + "`r$($newLine)"
      Add-RavelloApplicationVm @sAddVm | Set-RavelloApplicationDocumentation -Documentation $doc -Confirm:$false | Out-Null

      # Adjust VM
      if($_.NumCpu -or $_.MemorySize -or $_.Hostname)
      {
        $sVM = @{
          ApplicationId = $app.id
          VmName = $_.VmName
          Confirm = $false
        }
        if($_.NumCpu)
        {
          $sVM.Add('NumCpu',$_.NumCpu)
        }
        if($_.MemorySize)
        {
          $sVM.Add('MemorySize',$_.MemorySize)
          $sVM.Add('MemorySizeUnit',$_.MemoryUnit)
        }
        if($_.Hostname)
        {
          $sVM.Add('Hostnames',$_.Hostname)
        }
        Set-RavelloApplicationVm @sVM | Out-Null
      }
      
      # Additional HD on VM
      if($_.HD)
      {
        foreach($hd in $_.HD){
          $sVM = @{
            ApplicationId = $app.id
            VmName = $_.VmName
            Confirm = $false
            DiskSize = $hd.HDSize
            DiskSizeUnit = $hd.HDUnit
          }
          Set-RavelloApplicationVmDisk @sVM | Out-Null
        }
      }
      
      # Add/Remove Services
      if($_.RDP -or $_.SSH)
      {
        $sVM = @{
          ApplicationId = $app.id
          VmName = $_.VmName
          Confirm = $false
        }
        if($_.RDP)
        {$sVM.Add('Rdp',[Convert]::ToBoolean($_.RDP))}
        if($_.SSH)
        {$sVM.Add('Ssh',[Convert]::ToBoolean($_.SSH))}
        Set-RavelloApplicationVmService @sVM | Out-Null
      }
            
      # Connect the ISO
      $sIso = @{
        ApplicationId = $app.id
        VmName = $_.VmName
        DiskImageName = &{
          foreach($iso in $lab.ISO){
            if($iso.Label -eq $_.ISO){$iso.FileName}
        }}
        Confirm = $false
      }
      Set-RavelloApplicationVmIso @sIso | Out-Null

      # Add to Order table
      $order += New-Object PSObject -Property @{
        Group = $_.Order
        VM = $_.VmName
        Delay = $_.OrderDelay
      }

      Write-Host -ForegroundColor Green 'done!'
    }

    # Publish the Application
    if(!(Test-RavelloApplicationPublished -ApplicationName $_.Labname))
    {
      Publish-RavelloApplication -ApplicationName $_.Labname -OptimizationLevel PERFORMANCE_OPTIMIZED -Confirm:$false | Out-Null 
    }
      
    # Set the start order
    $order = $order | Group-Object -Property Group | Sort-Object -Property Name | %{
      New-Object PSObject -Property @{
        Name = "Group$($_.Group[0].Group)"
        DelaySeconds = $_.Group[0].Delay
        VM = $_.Group | Select-Object -ExpandProperty VM
      }
    }
    New-RavelloApplicationOrderGroup -ApplicationId $app.id -StartOrder $order | Out-Null
    Write-Host 'Done'
  }
  else
  {
    # Warning when no Force
    Write-Host 'Application ' -NoNewline
    Write-Host -ForegroundColor Green "$($_.LabName)" -NoNewline
    Write-Host ' already exists.'
    Write-Host 'Use the ' -NoNewline
    Write-Host -ForegroundColor Red 'Force' -NoNewline
    Write-Host ' option to overwrite!' 
  }
}

Get-Date -Format 'hh:mm:ss'
#endregion

#region Stop
Disconnect-Ravello -Confirm:$false
#endregion
