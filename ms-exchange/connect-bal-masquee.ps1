<#
Auteur: @xhark http://blogmotion.fr
But:  Afficher le chemin d'une BAL exchange
Licence: Creative Commons BY-NC-SA 4.0
#>

param ( 
	[parameter(mandatory=$true, ValueFromPipeline=$true)]
	[string]$mailbox
)

### VARIABLES
$version="v2017.07.18"
$domain="masociete.local"

###############################################################################################

### ELEMENTS FENETRES
$host.ui.RawUI.WindowTitle = "Connexion d'une BAL masquee - " + $version
$pshost = Get-Host; 
$pswindow = $pshost.ui.rawui
$newsize = $pswindow.buffersize
$newsize.width = 250
$pswindow.buffersize = $newsize

# chargement module
if(([AppDomain]::CurrentDomain.GetAssemblies() | %{$_.ManifestModule.name}) -notcontains "System.DirectoryServices.AccountManagement.dll") {
		Add-Type -AssemblyName System.DirectoryServices.AccountManagement
		if(([AppDomain]::CurrentDomain.GetAssemblies() | %{$_.ManifestModule.name}) -notcontains "System.DirectoryServices.AccountManagement.dll") {
			Write-Host -Foregroundcolor Red "`r`nERREUR : IMPOSSIBLE DE CHARGER LE MODULE`r`n"
			exit
		}
}

# contexte
$ct=[System.DirectoryServices.AccountManagement.ContextType]::Domain
$user=[System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ct,"$domain\$mailbox")

if ($user)	{
	$result=$user.GetUnderlyingObject().legacyExchangeDN
	Write-Host -ForegroundColor Yellow "`r`nVoici le chemin de la boite pour le compte AD '$mailbox' :"
	Write-Host -ForegroundColor Green "$result"
	"$result" | clip
	Write-Host "`r`n"
} else {
	Write-Host -ForegroundColor Red "`r`nERREUR : Aucune BAL pour le compte AD '$mailbox'r`n"
}


