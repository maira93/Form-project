<#
Requisitos:
- PowerShell 7+
- Módulo Microsoft.Graph (Install-Module Microsoft.Graph -Scope CurrentUser)
- Permissões: User.ReadWrite.All, Group.ReadWrite.All

Uso:
pwsh ./CriarUsuariosEquipeTeste.ps1 -Domain "contoso.onmicrosoft.com" -Count 10 -GroupDisplayName "EquipeTeste"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,
    [int]$Count = 10,
    [string]$GroupDisplayName = "EquipeTeste",
    [string]$OutputCsvPath
)

# Fail fast
$ErrorActionPreference = "Stop"

# Define caminho padrão para CSV (se não informado)
if (-not $OutputCsvPath -or [string]::IsNullOrWhiteSpace($OutputCsvPath)) {
    $defaultName = "usuarios_criados_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    $OutputCsvPath = Join-Path -Path (Get-Location) -ChildPath $defaultName
}

# Conecta no Graph se ainda não conectado
try {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Import-Module Microsoft.Graph -ErrorAction Stop
        Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All" | Out-Null
    }
} catch {
    throw "Falha ao conectar no Microsoft Graph. Detalhes: $($_.Exception.Message)"
}

# Dados fictícios
$firstNames = @(
    "Ana","Bruno","Carla","Diego","Eduarda","Felipe","Gabriela","Hugo","Isabela","João",
    "Kaio","Larissa","Marcos","Natália","Otávio","Paula","Rafael","Sofia","Tiago","Valentina",
    "Wagner","Yasmin","Zeca","Beatriz","Clara","Daniel","Elisa","Gustavo","Henrique","Iara"
)
$lastNames = @(
    "Silva","Santos","Oliveira","Souza","Rodrigues","Ferreira","Almeida","Costa","Gomes","Martins",
    "Araujo","Ribeiro","Carvalho","Lopes","Barbosa","Melo","Castro","Cardoso","Correia","Dias",
    "Teixeira","Fernandes","Cavalcante","Pereira","Rocha","Miranda","Farias","Freitas","Machado","Nogueira"
)
$funTitles = @(
    "Especialista em memes","Vice-presidente de café","Ninja de planilhas","Guardião das senhas",
    "Administrador de caos","Evangelista de GIFs","Curador de playlists","CEO de Atalhos",
    "Sommelier de bugs","Arquiteto de gambiarra","Piloto de Deploys","Mestre dos tickets",
    "Engenheiro de Café com Leite","Healer de Incidentes","Cientista de Happy Hour"
)

function New-RandomPassword {
    param(
        [int]$Length = 14,
        [int]$NumNonAlphanumeric = 3
    )
    # Compatível com PowerShell 7 (sem System.Web)
    $lower = 'abcdefghijklmnopqrstuvwxyz'
    $upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $special = '!@#$%^&*()-_=+[]{}:,.?'
    $all = ($lower + $upper + $digits + $special)

    $chars = New-Object System.Collections.Generic.List[char]
    # Garante complexidade mínima
    $chars.Add(($upper.ToCharArray() | Get-Random))
    $chars.Add(($lower.ToCharArray() | Get-Random))
    $chars.Add(($digits.ToCharArray() | Get-Random))
    for ($i = 0; $i -lt $NumNonAlphanumeric; $i++) { $chars.Add(($special.ToCharArray() | Get-Random)) }
    while ($chars.Count -lt $Length) { $chars.Add(($all.ToCharArray() | Get-Random)) }
    $chars = $chars | Sort-Object { Get-Random }
    -join $chars
}

function New-UserNameCombo {
    param([string]$Domain)
    $fn = Get-Random -InputObject $firstNames
    $ln = Get-Random -InputObject $lastNames
    $mailNick = ($fn + $ln) -replace '[^a-zA-Z0-9]', '' | ForEach-Object { $_.ToLower() }
    $upn = ("{0}.{1}@{2}" -f $fn,$ln,$Domain).ToLower()
    $display = "$fn $ln"
    [pscustomobject]@{
        GivenName         = $fn
        Surname           = $ln
        MailNickname      = $mailNick
        UserPrincipalName = $upn
        DisplayName       = $display
    }
}

# Garante o grupo "EquipeTeste"
function Ensure-Group {
    param([string]$DisplayName)
    $escaped = $DisplayName.Replace("'","''")
    $grp = Get-MgGroup -Filter "displayName eq '$escaped'" -All -ConsistencyLevel eventual | Select-Object -First 1
    if (-not $grp) {
        # mailNickname precisa ser único (sem espaços e minúsculo)
        $mailNick = ($DisplayName -replace '\\s','') .ToLower()
        # Se já existir apelido igual, anexa um sufixo aleatório
        $existsNick = Get-MgGroup -Filter "mailNickname eq '$mailNick'" -All -ConsistencyLevel eventual | Select-Object -First 1
        if ($existsNick) { $mailNick = "$mailNick$([System.Guid]::NewGuid().ToString('N').Substring(0,6))" }

        $grp = New-MgGroup -DisplayName $DisplayName `
                           -MailEnabled:$false `
                           -MailNickname $mailNick `
                           -SecurityEnabled:$true `
                           -Description "Grupo criado automaticamente pelo script em $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        Write-Host "Grupo criado: $($grp.DisplayName) ($($grp.Id))"
    } else {
        Write-Host "Grupo já existe: $($grp.DisplayName) ($($grp.Id))"
    }
    return $grp
}

$group = Ensure-Group -DisplayName $GroupDisplayName
$createdUsers = @()
$seenUpns = New-Object System.Collections.Generic.HashSet[string]

for ($i = 1; $i -le $Count; $i++) {
    # Gera uma combinação de nome/UPN (evita duplicados locais na mesma execução)
    $combo = $null
    do {
        $combo = New-UserNameCombo -Domain $Domain
    } while (-not $seenUpns.Add($combo.UserPrincipalName))

    # Verifica existência
    $upnEscaped = $combo.UserPrincipalName.Replace("'","''")
    $existing = Get-MgUser -Filter "userPrincipalName eq '$upnEscaped'" -All -ConsistencyLevel eventual | Select-Object -First 1

    if ($existing) {
        Write-Host "Usuário já existe, não será sobrescrito: $($existing.UserPrincipalName)"
        $userToUse = $existing
    } else {
        # Garante mailNickname único
        $nickEscaped = $combo.MailNickname.Replace("'","''")
        $nickExists = Get-MgUser -Filter "mailNickname eq '$nickEscaped'" -All -ConsistencyLevel eventual | Select-Object -First 1
        if ($nickExists) {
            $combo.MailNickname = "$(
                $combo.MailNickname
            )$([System.Guid]::NewGuid().ToString('N').Substring(0,4))"
        }

        $pwd = New-RandomPassword
        $jobTitle = Get-Random -InputObject $funTitles

        try {
            $newUser = New-MgUser `
                -AccountEnabled:$true `
                -DisplayName $combo.DisplayName `
                -GivenName $combo.GivenName `
                -Surname $combo.Surname `
                -UserPrincipalName $combo.UserPrincipalName `
                -MailNickname $combo.MailNickname `
                -UsageLocation "BR" `
                -JobTitle $jobTitle `
                -PasswordProfile @{
                    forceChangePasswordNextSignIn = $true
                    password = $pwd
                }

            Write-Host "Usuário criado: $($newUser.UserPrincipalName) | Senha temporária: $pwd | JobTitle: $jobTitle"
            $userToUse = $newUser
            $createdUsers += [pscustomobject]@{
                UserPrincipalName = $newUser.UserPrincipalName
                DisplayName       = $combo.DisplayName
                GivenName         = $combo.GivenName
                Surname           = $combo.Surname
                MailNickname      = $combo.MailNickname
                UsageLocation     = 'BR'
                TemporaryPassword = $pwd
                JobTitle          = $jobTitle
                Id                = $newUser.Id
            }
        } catch {
            Write-Warning "Falha ao criar $($combo.UserPrincipalName): $($_.Exception.Message)"
            continue
        }
    }

    # Adiciona ao grupo (existente ou recém-criado)
    try {
        New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($userToUse.Id)"
        } | Out-Null
        Write-Host "Adicionado ao grupo '$($group.DisplayName)': $($userToUse.UserPrincipalName)"
    } catch {
        if ($_.Exception.Message -match 'One or more added object references already exist') {
            Write-Host "Já é membro do grupo: $($userToUse.UserPrincipalName)"
        } else {
            Write-Warning "Falha ao adicionar ao grupo: $($userToUse.UserPrincipalName) - $($_.Exception.Message)"
        }
    }
}

# Saída resumida (somente criados agora) + exportação CSV
if ($createdUsers.Count -gt 0) {
    "`nResumo dos usuários criados:" | Write-Host
    $createdUsers | Select-Object UserPrincipalName, DisplayName, GivenName, Surname, MailNickname, UsageLocation, TemporaryPassword, JobTitle | Format-Table -AutoSize | Out-Host

    try {
        $createdUsers | Select-Object UserPrincipalName, DisplayName, GivenName, Surname, MailNickname, UsageLocation, TemporaryPassword, JobTitle |
            Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Arquivo CSV salvo em: $OutputCsvPath"
    } catch {
        Write-Warning "Falha ao exportar CSV para '$OutputCsvPath': $($_.Exception.Message)"
    }
} else {
    Write-Host "`nNenhum novo usuário foi criado nesta execução."
}