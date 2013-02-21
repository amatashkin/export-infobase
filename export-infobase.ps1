
#  export-infobase
#  Powershell script for exporting 1C 8.2 infobase to .dt file.
#  Экспорт информационной базы в DT
#  * блокирует соединений с ИБ
#  * разрывает активные сессии в ИБ

# Init
$Cluster = ""
$Clusters = ""
$InfoBase = ""
$InfoBases = ""
$ServerAgent = ""
$V82Com = ""
$ClusterFound = $False
$InfoBaseFound = $False

# Указываем путь к 1С и проверяем наличие.
$str1CPath='C:\Program Files (x86)\1cv82\common\1cestart.exe'
if (!(Test-Path $str1CPath)) { "1C is missing at $str1CPath" ; exit 13 }

# Параметры запуска: адрес сервера, основной порт кластера, информационная база
$hostname = $env:computername
$ServerAddress = 'tcp://' + $hostname + ':1540'
$MainPort = "1541"
$InfoBaseName = "UT"

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
        break
        # Пишем в лог о находке
        }
    }

# Проверка кластера
if (!($ClusterFound))
    {
     # Пишем в лог что кластер не найден
     write-host "Не найден кластер серверов 1С"
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
    if (!($WorkingProcess.Running -eq 1) )
        {
        continue   
        }
     
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

# Проверка базы
if (!($InfoBaseFound))
    {
    write-host "Не найдена указанная информационная база."
    # пишем в лог
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
write-host "Найдено сессий: " $Sessions.Count
$Sessions | ft

# Завершаем сессисии
foreach ($Session in $Sessions)
    {
    $ServerAgent.TerminateSession($Cluster,$Session)
    }

# Проверяем что все отключилось
$Sessions = $ServerAgent.GetInfoBaseSessions($Cluster, $Base)
if (!($Sessions.Count -eq 0))
    {
    write-host "Не удалость отключить сессий:" $Sessions.Count
    $Sessions | ft
    # Ошибка. Пишем в лог. Не удалость отключить часть сесссий
    }

# ===============================================
# 
# Пробуем выгрузить .DT
write-host "Выгрузка ИБ..."
start powershell -ArgumentList "C:\temp\save-DT.ps1 test rutest01;exit $LASTEXITCODE"
# 
# ===============================================

# Снятие блокировки соединений ИБ и кода доступа
$InfoBase.ConnectDenied = $False
$InfoBase.PermissionCode = ""
$CWP.UpdateInfoBase($InfoBase)

$V82Com = ""