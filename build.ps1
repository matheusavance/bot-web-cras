$exclude = @("venv", "teste_bot.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "teste_bot.zip" -Force