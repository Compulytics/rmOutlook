$SystemDrive = (Get-WmiObject Win32_OperatingSystem).SystemDrive
get-childitem -path $SystemDrive\Users | Select Name | Foreach-Object {
	if(($_.Name -ne "Public") -and !($_.Directory)){
		$UserName = $_.Name
		Remove-Item $SystemDrive\Users\$UserName\AppData\Local\Microsoft\Outlook\*.* -force -ErrorAction SilentlyContinue
		sleep 5
		if ($UserName -ne $Env:UserName){
			reg load HKLM\User-Hive $SystemDrive\Users\$UserName\ntuser.dat
			Remove-ItemProperty -Path HKLM:\User-Hive\Software\Microsoft\Office\16.0\Outlook\Setup\ -name First-Run -ErrorAction SilentlyContinue
			Remove-Item HKLM:\User-Hive\Software\Microsoft\Office\16.0\Outlook\Profiles\outlook -Recurse -Force -ErrorAction SilentlyContinue
			[gc]::collect()
			sleep 5
			reg unload HKLM\User-Hive
		}
		else{
			Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup\ -name First-Run -ErrorAction SilentlyContinue
			Remove-Item HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\outlook -Recurse -Force -ErrorAction SilentlyContinue
		}
	}
}
