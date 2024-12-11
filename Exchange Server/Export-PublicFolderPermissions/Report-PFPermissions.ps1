# Copyright Frank Carius
# https://www.msxfaq.de/exchange/tools/auswertung/reportpfpermissions.htm

Begin {
} 

Process {
	$pf = $_
	if($pf.parentpath -eq '\') {
		$folder = $pf.parentpath + $pf.name
	}
	else {
		$folder = $pf.parentpath + "\" + $pf.name
	}
	write-host  "Processing Folder:" $folder 
	$permission = get-PublicFolderClientPermission -identity $pf.Identity
	foreach ($perm in $permission) {
		$rights = [string]$perm.accessRights 
		$User = [string]$perm.User 
#			write-host "Rechte" $rights "User:" $User
		$pso = New-Object PSObject
		Add-Member -InputObject $pso noteproperty Folder $folder
		Add-Member -InputObject $pso noteproperty User   $User
		Add-Member -InputObject $pso noteproperty AccessRights $rights
		$pso 
	}
}

End {
}