<#
#
# Script criado por Alan Godoi da Silveira
# 17/05/2016
#
# Requisitos:
#
# 1- É necessário executar o script com um usuário (Windows) que possua acesso à toda rede.
# Também é possível executar o script com um usuário que seja administrador nos computadores.
#
# 2- Os computadores clientes precisam ter o serviço RemoteRegistry rodando.
#>

#$Array = @() ## Create Array to hold the Data
<##
#    Arquivo com os nomes dos computadores
#    
#    Exemplo:
#    E001
#    E002
#    E003
#    E004
#    E005
#
#    $Computers recebe os os nomes dos computadores
##>
$Computers = Get-Content -Path C:\scripts\pc-names.txt

<##
#    Loop para cada linha no arquivo $Computers (C:\scripts\pc-name.txt)
##>
foreach ($Computer in $Computers)
{
    <##
    #    Teste de ping
    ##>
    if (Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue){
        
        <##
        #    $DSC = Descrição do Computador - Windows + Pause/Break
        #    $OS = Sistema Operacional
        #    $CPU = Processador
        #    $RAM = Tamanho da memória RAM - Físico
        #    $HDD = Tamanho do disco rígido - Físico
        #    $USR = Último usuário logado
        ##>
        $DSC = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer | ForEach-Object -MemberName Description
        $OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer | ForEach-Object -MemberName Caption
        $CPU = Get-WmiObject win32_processor -ComputerName $Computer | ForEach-Object -MemberName Name
        $RAM = Get-WMIObject -class Win32_PhysicalMemory -ComputerName $Computer | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}
        $HDD = Get-WMIObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $Computer | ForEach-Object {[math]::truncate($_.size / 1GB)}
        $USR = (Get-WMIObject -class Win32_ComputerSystem -ComputerName $Computer | select username).username

        <#GET OFFICE
        #
        #    https://social.technet.microsoft.com/Forums/scriptcenter/en-US/927049c9-25d0-4641-bd44-865e9750e514/powershell-script-to-find-office-version?forum=ITCG
        #    Precisa ter o serviço de acesso remoto ao registro do Windows (RemoteRegistry) rodando nos computadores
        #
        #    Rode o comando "Get-Service remoteregistry -ComputerName E001" para verificar se o serviço está rodando em um computador remoto.
        #
        #>
        $version = 0
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)

        $reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() |% {
            if ($_ -match '(\d+)\.') {
                if ([int]$matches[1] -gt $version) {
                    $version = $matches[1]
                }
            }    
    }

        #Change Office version to year
        if ($version.Equals("15")){
            $version = "2013"
        } elseif ($version.Equals("14")){
            $version = "2010"
        } elseif ($version.Equals("12")){
            $version = "2007"
        } elseif ($version.Equals("11")){
            $version = "2003"
        }

        Write-Host "Computador:" $Computer
        Write-Host "Setor:" $DSC
        Write-Host "SO:" $OS
        Write-Host "Processador:" $CPU
        Write-Host "Memória RAM (GB):" $RAM
        Write-Host "HD (GB):" $HDD
        Write-Host "Usuário:" $USR
        Write-Host "Office:" $version
        Write-Host -fore Red "======"

        echo $("Computador: " + $Computer) >> C:\scripts\teste.txt
        echo $("Setor: " + $DSC) >> C:\scripts\teste.txt
        echo $("SO: " + $OS) >> C:\scripts\teste.txt
        echo $("Processador: " + $CPU) >> C:\scripts\teste.txt
        echo $("Memória RAM (GB): " + $RAM) >> C:\scripts\teste.txt
        echo $("HD (GB): " + $HDD) >> C:\scripts\teste.txt
        echo $("Usuário: " + $USR) >> C:\scripts\teste.txt
        echo "======" >> C:\scripts\teste.txt

    }
    else{
        #Write-Host "$name,down"
        Write-Host "O computador"$Computer "está Off-line"
        Write-Host "======"
    }
}

    #$Result = "" | Select PCName,OS_Name, Processador, RAM ## Create Object to hold the data
    
    #$OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer
    ##$OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer | ForEach-Object -MemberName Caption
    
    ##$CPU = Get-WmiObject win32_processor -ComputerName $Computer | ForEach-Object -MemberName Name
    
    #$RAMB = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer | ForEach-Object -MemberName TotalPhysicalMemory
    #$RAM = Get-WMIObject -Class Win32_ComputerSystem  -ComputerName $Computer | ForEach-Object {[math]::truncate($_.TotalPhysicalMemory / 1GB)}
    ##$RAM = Get-WMIObject -class Win32_PhysicalMemory -ComputerName $Computer | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}


    #$HDDB = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $Computer | ForEach-Object -MemberName Size
    ##$HDD = Get-WMIObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $Computer | ForEach-Object {[math]::truncate($_.size / 1GB)}
    
    ##$USR = (Get-WMIObject -class Win32_ComputerSystem -ComputerName $Computer | select username).username

    

