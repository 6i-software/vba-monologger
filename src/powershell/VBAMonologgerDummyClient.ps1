# To enable CLI option -v/-Verbose (and more...) 
[CmdletBinding()]
param (
    [Parameter(Position = 0)][string] $message = "Dummy message to send to VBAMonologger server"
)

Add-Type -AssemblyName System.Net.Http

# Définir l'URL du serveur HTTP
$uri = "http://localhost:20100"

# Créer un objet HTTP client
$client = New-Object System.Net.Http.HttpClient

# Créer un message HTTP de type POST avec le contenu "exit"
$content = New-Object System.Net.Http.StringContent($message, [System.Text.Encoding]::UTF8, "text/plain")

# Envoyer la requête POST
$response = $client.PostAsync($uri, $content).Result

# Afficher le statut de la réponse et le contenu
[console]::WriteLine("HTTP response: $($response.StatusCode)")
$responseContent = $response.Content.ReadAsStringAsync().Result
[console]::WriteLine("Body response: $responseContent")

# Fermer le client HTTP
$client.Dispose()