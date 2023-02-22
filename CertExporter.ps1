Function BrowseForFiles($initialDirectory, [switch]$SaveAs,[switch]$tsharkLoc)
{   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	if ($SaveAs) {
		$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
	} else {
		$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	}
	
	$OpenFileDialog.initialDirectory = $initialDirectory
	if ($SaveAs) {
        $OpenFileDialog.title = "Save As Cert File"
		$OpenFileDialog.filter = "All files (*.cer)| *.cer"
    }
    else {
        $OpenFileDialog.title = "Open File pcap file"
        $OpenFileDialog.filter = "Pcap file (*.pcap)| *.pca*"
    }
	$OpenFileDialog.filename = $DefFiles
	$OpenFileDialog.ShowDialog() | Out-Null
	$FileName = $OpenFileDialog.filename
	$FileName
} #end function BrowseForFiles

Function Get-Folder($initialDirectory) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'
    if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
    [void] $FolderBrowserDialog.ShowDialog()
    return $FolderBrowserDialog.SelectedPath
}

if (Test-Path "c:\program Files\Wireshark\tshark.exe") {$PcapCert = (&"C:\Program Files\Wireshark\tshark.exe" -n -Tfields -e tls.handshake.certificate -Y tls.handshake.certificate -r (BrowseForFiles -initialDirectory "C:\temp\") )}`
 elseif (Test-Path ".\tshark.exe") { $PcapCert = (&".\tshark.exe" -n -Tfields -e tls.handshake.certificate -Y tls.handshake.certificate -r (BrowseForFiles -initialDirectory "C:\temp\") )} `
 elseif (Test-Path "c:\program Files (x86)\Wireshark\tshark.exe") { $PcapCert = (&"c:\program Files (x86)\Wireshark\tshark.exe" -n -Tfields -e tls.handshake.certificate -Y tls.handshake.certificate -r (BrowseForFiles -initialDirectory "C:\temp\") ) }`
 else { Write-Host "Error: You must have a tshark.exe file on your computer! {install Wireshark}" -ForegroundColor Red -BackgroundColor Black; Break} 
if ($PcapCert.length -lt 1) {Write-Host "Error: looks like the pcap file not have any certs" -ForegroundColor Red -BackgroundColor Black; Start-Sleep -Seconds 5 ;Break}
elseif ($PcapCert.length -eq 1) {
	[byte[]]$bytes = ($PcapCert[0]  -split '(.{2})' -ne '' -replace '^', '0X')
	$MyCert = "-----BEGIN CERTIFICATE-----`r`n{0}`r`n-----END CERTIFICATE-----" -f [System.Convert]::ToBase64String($bytes)
	[IO.File]::WriteAllLines((BrowseForFiles -SaveAs -initialDirectory "C:\temp\"), $MyCert)
} else {
	Write-Host "Looks like there is more then 1 certs in the pcap file, i'll generate names by numbers... please select path to save the certs..." -ForegroundColor Green
	($PcapCert.length)
	$myFol = (Get-Folder)
	for ($i=0; $i -le ($PcapCert.length-1); $i++) {
		[byte[]]$bytes = ($PcapCert[$i]  -split '(.{2})' -ne '' -replace '^', '0X')
		$MyCert = "-----BEGIN CERTIFICATE-----`r`n{0}`r`n-----END CERTIFICATE-----" -f [System.Convert]::ToBase64String($bytes)
		$FileToSave = "{0}\{1}.cer" -f $myFol,$i
		[IO.File]::WriteAllLines($FileToSave, $MyCert)
		$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $FileToSave
		$CertName = ((($cert.Subject | Select-String "CN=([a-z0-9|-]+\.)*[a-z0-9|-]+\.[a-z]+" -AllMatches).Matches).Value).Replace("CN=","")+".cer"
		Rename-Item $FileToSave $CertName -Confirm:$false -Force
	}
}

