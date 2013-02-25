export-infobase
===============

Powershell script for exporting 1C 8.2 infobase to .dt file.

Скрипт для выгрузки информационной базы 1С:Предприятие 8.2 в файл .DT
Использует менеджер COM соединений сервера приложений.

Пример использования:

PS> .\export-infobase.ps1 InfoBaseName

где InfoBaseName - это имя информационной базы на сервере 1С.

На данный момент пишет логи в C:\Logs а DT в C:\1C_Backup\$Year\$InfoBaseName

ToDo
====

[ ] Запись лога в EventLog