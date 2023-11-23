
new-adserviceaccount -Name 'GMSA-Test' -DNSHostName 'GMSA-Test.damborsky.com' -PrincipalsAllowedToRetrieveManagedPassword 'Domain Admins' -ServicePrincipalNames 'HTTP/GMSA-Test.domain.com'

new-adserviceaccount -Name 'GMSA-Test02' -DNSHostName 'GMSA-Test02.damborsky.com' -PrincipalsAllowedToRetrieveManagedPassword 'Domain Admins' -ServicePrincipalNames "$z440-3"

Get-ClusterSharedVolume | Get-SmbShare


