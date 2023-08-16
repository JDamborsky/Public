## Create regex to verify if it is a IP-adress
$regex = [regex]"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"

## Create a string with a IP-adress
$ip = "10.2.2.4"
## Check if the string is a IP-adress
if ($regex.IsMatch($ip)) {
    Write-Host "It is a IP-adress"
} else {
    Write-Host "It is not a IP-adress"
}


