# Deletes files and folers recursively that were last modified over 7 days ago

Get-ChildItem 'E:\RM' | Where-Object {$_.LastWriteTime -le (Get-Date).AddDays(-7)} | Foreach-Object { Remove-Item $_.FullName -Recurse -Force} 