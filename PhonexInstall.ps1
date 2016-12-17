    Function Uninstall-Program([string]$name) # http://thoughtsofmarcus.blogspot.ru/2012/12/clever-uninstall-of-msi.html
{
    $success = $false

    # Read installation information from the registry
    $registryLocation = Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
    foreach ($registryItem in $registryLocation)
    {
        # If we get a match on the application name
        if ((Get-itemproperty $registryItem.PSPath).DisplayName -eq $name)
        {
            # Get the product code if possible
            $productCode = (Get-itemproperty $registryItem.PSPath).ProductCode
            
            # If a product code is available, uninstall using it
            if ([string]::IsNullOrEmpty($productCode) -eq $false)
            {
                Write-Host "Uninstalling $name, ProductCode:$code"
            
                $args="/uninstall $code"

                [diagnostics.process]::start("msiexec", $args).WaitForExit()
                
                $success = $true
            }
            # If there is no product code, try to read the uninstall string
            else
            {
                $uninstallString = (Get-itemproperty $registryItem.PSPath).UninstallString
                
                if ([string]::IsNullOrEmpty($uninstallString) -eq $false)
                {
                    # Grab the product key and create an argument string
                    $match = [RegEx]::Match($uninstallString, "{.*?}")
                    $args = "/x $($match.Value) /qb"

                    [diagnostics.process]::start("msiexec", $args).WaitForExit()
                    
                    $success = $true
                }
                else { throw "Unable to uninstall $name" }
            }
        }
    }

    $registryLocation = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\"
    foreach ($registryItem in $registryLocation)
    {
        # If we get a match on the application name
        if ((Get-itemproperty $registryItem.PSPath).DisplayName -eq $name)
        {
            # Get the product code if possible
            $productCode = (Get-itemproperty $registryItem.PSPath).ProductCode
            
            # If a product code is available, uninstall using it
            if ([string]::IsNullOrEmpty($productCode) -eq $false)
            {
                Write-Host "Uninstalling $name, ProductCode:$code"
            
                $args="/uninstall $code"

                [diagnostics.process]::start("msiexec", $args).WaitForExit()
                
                $success = $true
            }
            # If there is no product code, try to read the uninstall string
            else
            {
                $uninstallString = (Get-itemproperty $registryItem.PSPath).UninstallString
                
                if ([string]::IsNullOrEmpty($uninstallString) -eq $false)
                {
                    # Grab the product key and create an argument string
                    $match = [RegEx]::Match($uninstallString, "{.*?}")
                    $args = "/x $($match.Value) /qb"

                    [diagnostics.process]::start("msiexec", $args).WaitForExit()
                    
                    $success = $true
                }
                else { echo "Unable to uninstall $name" }
            }
        }
    }


    
    if ($success -eq $false)
    { echo "Unable to find application $name" }
}

## set-executionpolicy remotesigned # to allow the scripts, run PS under admin
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") > $null

$modules = Get-ChildItem .\*.msi
# $regex = [regex] '([A-z]*)(|x86|x64)_.*' # no underscore - PhonexComparex64_20160623
$regex = [regex] '([A-z]*)(|_x86|_x64)_.*' # underscore - PhonexCompare_x86_20160625 PhonexCards_20160625
foreach ($m in $modules) 
    {
    if (!($m.name.Contains("Collect") `
        -or $m.name.Contains("Guests") `
        -or $m.name.Contains("NetflowDetails") `
        -or $m.name.Contains("Web") `
        -or $m.name.Contains("x64") `
		-or $m.name.Contains("Administrator") `
        -or $m.name.Contains("Statistics") `
		#-or $m.name.Contains("Cards") `
        )) 
        { 
        $appname = $regex.Replace($m.name, '$1')
		        if ( `
            $m.name.Contains("Hotel") `
            -or $m.name.Contains("Kernel") `
            -or $m.name.Contains("Radius") 
            )
            {
                Stop-Service -name $appname
            }
		#if ($m.name.Contains("Kernel")) {Uninstall-Program "PhonexImport"}
        Uninstall-Program $appname
        echo "$appname uninstalled"
		$retry = "Retry"
		while ($retry -eq "Retry")
        {
            if ( (Start-Process -FilePath "msiexec.exe" -ArgumentList ("/i " + $m.FullName + " ALLUSERS=""2"" /Qb") -Wait -Passthru -Verb runAs).ExitCode -ne 0) 
    		{ 
    			$retry = [System.Windows.Forms.MessageBox]::Show("Error installing $m.name","","AbortRetryIgnore","Warning")
    		}
    		else { 
                echo "$appname installed"
                $retry = "Abort" 
                }
        }

        if ( `
            $m.name.Contains("Hotel") `
            -or $m.name.Contains("Kernel") `
            -or $m.name.Contains("Radius") 
            )
            {
                Set-Service -name $appname -StartupType Manual
                # Start-Process -FilePath sc.exe -ArgumentList ("config " + $appname + " obj= SERVICE_ACCOUNT_NAME_HERE password= SERVICE_PASSWORD_HERE") -Verb runAs
            }
        }
        
    }
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")