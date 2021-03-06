﻿
#  export-infobase
#  Powershell script for exporting 1C 8.2 infobase to .dt file.
#  Экспорт информационной базы в DT
#  * блокирует соединений с ИБ
#  * разрывает активные сессии в ИБ

param([string]$InfoBaseName, [string]$ServerName = "$Env:Computername", [string]$strLogPath = 'C:\Logs')


function Write-LogFile([string]$logFileName)
{
    Process
    {
        $_
        $dt = Get-Date
        $str = $dt.DateTime + " " + $_
        $str | Out-File -FilePath $logFileName -Append
    }
}

function Start-ProcessTree([string]$FilePath = $(Read-Host "Supply a value for the FilePath parameter"),[string]$ArgumentList,[int]$TimeoutMin = 60,[switch]$WaitForChildProcesses)
{
    $Timeoutms = $TimeoutMin * 60 * 1000
    $ProcessStartInfo = New-Object System.Diagnostics.ProcessStartInfo $FilePath
    if ($ArgumentList) { $ProcessStartInfo.Arguments = $ArgumentList }
    $Process = [System.Diagnostics.Process]::Start($ProcessStartInfo)
    $ProcessId = $Process.Id
    $ProcessStartTime = $Process.StartTime
    $ProcessCompleted = $Process.WaitForExit($Timeoutms)
    while ($WaitForChildProcesses -and $ProcessCompleted)
    {
      [array]$ChildProcesses = Get-WmiObject Win32_Process -Filter "ParentProcessId = $ProcessId"
      if (!$ChildProcesses.Count)
      {
        break
      }
      $Elapsedms = (New-TimeSpan $ProcessStartTime (Get-Date)).TotalMilliseconds
      if ($Elapsedms -lt $Timeoutms)
      {
        Start-Sleep -Seconds 1
      }
      else
      {
        $ProcessCompleted = $false
      }
    }
    return $ProcessCompleted
}

# Init
$Cluster = ""
$Clusters = ""
$InfoBase = ""
$InfoBases = ""
$ServerAgent = ""
$V82Com = ""
$ClusterFound = $False
$InfoBaseFound = $False

# Проверяем входные параметры. Если нету - убегаем.
if (!$InfoBaseName) { "Database name is missing."; exit 11}
if (!$ServerName) { "Servername name is missing."; exit 12}

# Указываем путь к 1С и проверяем наличие.
$str1CPath='C:\Program Files (x86)\1cv82\common\1cestart.exe'
if (!(Test-Path $str1CPath)) { "1C is missing at $str1CPath" ; exit 13 }

# Параметры запуска: адрес сервера, основной порт кластера, информационная база
$ServerAddress = 'tcp://' + $ServerName + ':1540'
$MainPort = "1541"

# Устанавливаем переменные для дат
$StartYear = Get-Date -uFormat %Y
$TimeStamp  = Get-Date -uFormat %H%M%S
$StartDate = Get-Date -uFormat %Y-%m-%d
$StartTime = Get-Date -uFormat %H:%M:%S

# Проверяем наличие директории для логов и бекапов.
# Создаем всякие пути.
if (!(test-path $strLogPath)) {new-item $strLogPath -type directory}
$strLogName = $strLogPath + '\' + 'export-infobase.' + $InfoBaseName + '.log'
$strDBPath = "$ServerName\$InfoBaseName"
$strBackupPath = "C:\1C_Backup\$StartYear\$InfoBaseName"
if (!(test-path $strBackupPath)) {new-item $strBackupPath -type directory}
$strBackupName = "$strBackupPath\$InfoBaseName`_$StartDate`_$TimeStamp.dt"
# Параметры запуска для 1С:
$arguments1C = "CONFIG /S `"$strDBPath`" /DisableStartupMessages /DumpIB `"$strBackupName`" /Out `"$strLogName`" -NoTruncate /UC `"$StartYear`" "

"==========================================" | Write-LogFile $strLogName
"Job started $StartDate at $StartTime" | Write-LogFile $strLogName

$V82Com = New-Object -COMObject "V82.COMConnector"

# Подключение к агенту сервера
$ServerAgent = $V82Com.ConnectAgent($ServerAddress)

# Поиск нужного кластера
$Clusters = $ServerAgent.GetClusters()
foreach ($Cluster in $Clusters)
{
    if ($Cluster.MainPort -eq $MainPort)
    {
        $ClusterFound = $True    
        # Пишем в лог о находке
        "Найден кластер серверов: " + $Cluster.Name + " на " + $Cluster.HostName | Write-LogFile $strLogName
        break
    }
}

# Проверка кластера
if (!($ClusterFound))
{
     # Пишем в лог что кластер не найден
     "Не найден кластер серверов 1С" | Write-LogFile $strLogName
     break
}

# Аутентификация к выбранному кластеру
# если у пользователя под которым будет выполняться сценарий нет прав на кластер,
# можно прописать ниже имя пользователя и пароль администратора кластера
$ServerAgent.Authenticate($Cluster,"","")

# Получение списка рабочих процессов кластера
$WorkingProcesses = $ServerAgent.GetWorkingProcesses($Cluster)

# Поиск нужной базы
foreach ($WorkingProcess in $WorkingProcesses)
{
    if ($WorkingProcess.Running -eq 1)
    {
        $CWPAddr = "tcp://"+$WorkingProcess.HostName+":"+$WorkingProcess.MainPort
        $CWP= $V82Com.ConnectWorkingProcess($CWPAddr)
        $CWP.AddAuthentication("", "")

        $InfoBases = $CWP.GetInfoBases()

        foreach ($InfoBase in $InfoBases)
        {
            if ($InfoBase.Name -eq $InfoBaseName )
            {
                $InfoBaseFound = $True
                break
            }
        }

        if ($InfoBaseFound)
        {
            break
        }

    }
}

# Проверка базы
if (!($InfoBaseFound))
{
    # пишем в лог
    'Не найдена указанная информационная база.' | Write-LogFile $strLogName
    break
}

# Установка блокировки соединений ИБ и кода доступа (текущий год)
$InfoBase.ConnectDenied = $True
$InfoBase.PermissionCode = (Get-Date).Year
$CWP.UpdateInfoBase($InfoBase)

# Выбираем базу и подключеннные сессии
$Base = $ServerAgent.GetInfoBases($Cluster) | ? {$_.Name -eq $InfoBaseName}
$Sessions = $ServerAgent.GetInfoBaseSessions($Cluster, $Base)

# Пишем в лог количество сессий
"Останавливаем сессий: " + $Sessions.Count | Write-LogFile $strLogName

# Завершаем сессисии
foreach ($Session in $Sessions)
{
    "Отключаем: " + $Session.username + " на " + $Session.host + " начало работы " + $Session.startedat.datetime + " приложение " + $Session.appid | Write-LogFile $strLogName
    $ServerAgent.TerminateSession($Cluster,$Session)
}

# Проверяем что все отключилось
$Sessions = $ServerAgent.GetInfoBaseSessions($Cluster, $Base)
if (!($Sessions.Count -eq 0))
{
    # Ошибка. Пишем в лог. Не удалость отключить часть сесссий
    "Не удалость отключить сессий: " + $Sessions.Count | Write-LogFile $strLogName
    foreach ($Session in $Sessions)
    {
        "Не удалось отключить: " + $Session.username + " на " + $Session.host + " начало работы " + $Session.startedat.datetime + " приложение " + $Session.appid | Write-LogFile $strLogName
    }

}

# ===============================================
# 
# Пробуем выгрузить .DT
"Trying to backup $InfoBaseName at $ServerName" | Write-LogFile $strLogName
"Starting $str1CPath" | Write-LogFile $strLogName
"with parameters: $arguments1C" | Write-LogFile $strLogName

# Starting with default timeout 60 minutes
$Finished = Start-ProcessTree $str1CPath $arguments1C -WaitForChildProcesses 

$EndDate = Get-Date -uFormat %Y-%m-%d
$EndTime = Get-Date -uFormat %H:%M:%S
if ($Finished)
{
    "Success. Job finished $EndDate at $EndTime" | Write-LogFile $strLogName
    # Снятие блокировки соединений ИБ и кода доступа
    $InfoBase.ConnectDenied = $False
    $InfoBase.PermissionCode = ""
    $CWP.UpdateInfoBase($InfoBase)
    "Разблокировали базу" | Write-LogFile $strLogName
    "==========================================" | Write-LogFile $strLogName
}
else
{
    "Error. Job stopped by timeout $EndDate at $EndTime" | Write-LogFile $strLogName
    $InfoBase.ConnectDenied = $False
    $InfoBase.PermissionCode = ""
    $CWP.UpdateInfoBase($InfoBase)
    "Разблокировали базу" | Write-LogFile $strLogName
    "==========================================" | Write-LogFile $strLogName
    exit 10
}

# 
# ===============================================



$V82Com = ""