Clear-Host

# Installer Settings
$InstallFolder = 'C:\Program Files\Zabbix Agent 2'
$ZabbixMSIPath = 'E:\Всякая Куйня 2\zabbix_agent2-5.4.9-windows-amd64-openssl.msi'
if (!(Test-Path $ZabbixMSIPath))
{
    Write-Host "MSI пакет агента не обнаружен по этому пути: $ZabbixMSIPath" -ForegroundColor Red
    Break
}
$InstallerLogPath = Join-Path $env:TEMP ('zabbix_agent2_installer_' + (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") + '.log')

# Config settings
$LogType = 'file'
$LogFile = Join-Path $InstallFolder 'zabbix_agent2.log'
$Server = 'vds.lvovpd.ru'
$ServerActive = $Server
$Timeout = 15
$HostName = $env:COMPUTERNAME
$TlsConnect = 'psk'
$TlsAccept = 'psk'
$TlsPskIdentity = Read-Host -Prompt "Имя PSK ключа"
$TlsPskValue = Read-Host -Prompt "PSK ключ" -AsSecureString
$EnablePath = 1
$EnablePsk = 1

# Output Config
Clear-Host
$ConfigResults = @()
$HashTable = [ordered]@{
    'Путь установки агента' = $InstallFolder
    'Путь к логу установки' = $InstallerLogPath
    'Тип лога агента' = $LogType
    'Путь лога агента' = $LogFile
    'Сервер для пассивной проверки' = $Server
    'Сервер для активной проверки' = $ServerActive
    'Таймуат Агента' = $Timeout
    'Имя хоста' = $HostName
    'Включить PSK' = $EnablePsk
    'Протокол подключения' = $TlsConnect
    'Протокол ответа' = $TlsAccept
    'Имя PSK ключа' = $TlsPskIdentity
    'Ключ PSK' = if($TlsPskValue.Length -le 0) {'Ключ пустой'} else {'Ключ указан'}
    'Добавить в переменную PATH?' = $EnablePath
}
New-Object -TypeName PSObject -Property $HashTable

# CheckConfig
$Title    = 'Проверка настройки Агента'
$Question = 'Всё ли введено верно? Готовы продолжить?'

$Choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes','Продолжить тихую установку'))
$Choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No','Отменить установку и завершить работу скрипта'))

$Decision = $Host.UI.PromptForChoice($Title, $Question, $Choices, 1)
if ($decision -eq 0)
{
    Write-Host 'Запускаем установку' -ForegroundColor DarkGray

    # Запуск установки
    $ConvertSecureString = [System.Net.NetworkCredential]::new('', $TlsPskValue).Password

    $InstallExitCode = $null
    $InstallExitCode = Start-Process -FilePath msiexec -ArgumentList "/l*v `"$InstallerLogPath`" /i `"$ZabbixMSIPath`" /qn LOGTYPE=`"$LogType`" LOGFILE=`"$LogFile`" SERVER=`"$Server`" SERVERACTIVE=`"$ServerActive`" TIMEOUT=`"$Timeout`" HOSTNAME=`"$HostName`" TLSCONNECT=`"$TlsConnect`" TLSACCEPT=`"$TlsAccept`" TLSPSKIDENTITY=`"$TlsPskIdentity`" TLSPSKVALUE=`"$ConvertSecureString`" INSTALLFOLDER=`"$InstallFolder`" ENABLEPATH=`"$EnablePath`"" -Wait -PassThru

    if ($InstallExitCode.ExitCode -ge 1)
    {
        Write-Host ("Код выхода: " + $InstallExitCode.ExitCode + ". Код выхода не ноль, проверь код здесь: https://docs.microsoft.com/ru-ru/windows/win32/msi/error-codes") -ForegroundColor Red
    }
    else
    {
        $ServiceStatus = Get-Service 'Zabbix Agent 2'
        Write-Host ("Код выхода: " + $InstallExitCode.ExitCode + ". Установка успешна, проверяем службу...") -ForegroundColor Green
        Write-Host ("Статус службы: " + $ServiceStatus.Status + ", тип запуска: " + $ServiceStatus.StartType) -ForegroundColor Green
    }
}
else
{
    Write-Host 'Отмена. Скрипт завершён' -ForegroundColor Red
    Break
}