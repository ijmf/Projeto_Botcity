$exclude = @("venv", "Projeto_Botcity.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Projeto_Botcity.zip" -Force