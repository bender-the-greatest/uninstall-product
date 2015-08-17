#Functions

function wait-for-network ($tries){
   while(1){
      #Get a list of DHCP enabled interfaces that have a non-null DefaultIPGateway property.
      $x = gwmi -class Win32_NetworkAdapterConfiguration -filter DHCPEnabled=TRUE | where {$_.DefaultIPGateway -ne $null}
      
      #If there is at least one available, break the loop
      if (($x | measure).count -gt 0){break}
      
      #If $tries > 0 and we have tried $tries times without success, throw an exception
      if ($tries -gt 0 -and $try++ -ge $tries){
         throw "Network unavailable after $try tries."
      }
      
      #Wait
      start-sleep -s 60
   }
}

#End Functions

write-host "Finding COMODO Internet Security Product Code:`n"
$ProductCode = $null
$FoundComodo = $FALSE
#check every key in the uninstallation directory
Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -rec -ea SilentlyContinue | foreach {
   $UninstallKey = (Get-ItemProperty -Path $_.PsPath)
   
   #Check that the key has a DisplayName property, this is what we are checking
   if ($UninstallKey.DisplayName -ne $null){
   
      #check that the DisplayName property contains the common text
      if ($UninstallKey.DisplayName.ToLower().Contains("comodo internet security")){
      
         $FoundComodo=$TRUE
         write-host "Found MsiProductCode for" $UninstallKey.DisplayName
         
         #get the name of the key, must parse it from the full registry path
         $ProductCode = $_.PsPath.Split('\') | select -last 1
         
         #uninstall COMODO (COMODO messes with the count in the installer, so the /norestart option will not work)
         msiexec /X$ProductCode /qr
         break
      }
   }
}

#if it's not installed then write the confirmation file
if (!$FoundComodo){
   $Hostname=hostname
   
   #if the file exists then don't do anything
   if (!(test-path "\\danlawinc.com\Files\IT\.Comodo\$Hostname")){
      try{
         wait-for-network(5)
         
         #double check that the failure to find the file before wasn't b/c the network was down
         if (!(test-path "\\danlawinc.com\Files\IT\.Comodo\$Hostname")){
            echo $null > \\danlawinc.com\Files\IT\.comodo\$Hostname
         }
      } catch [System.Exception]{
      
         #load .net type System.Windows.Forms (we don't need it until here)
         [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
         [System.Windows.Forms.MessageBox]::Show("An error occurred informing Danlaw IT that Comodo is no longer installed on your system.`n`nPlease contact yoursupport@domain.tld as soon as possible regarding this error message.", "Autoremove Comodo");
      }
   }
}
