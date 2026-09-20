<#
.SYNOPSIS
    Identifica grupos dinâmicos do Microsoft Entra ID que usam a regra memberOf.

.DESCRIPTION
    Este script conecta ao Microsoft Graph e consulta todos os grupos dinâmicos
    do tenant. Em seguida, localiza regras de associação que utilizam memberOf
    ou user.memberOf, que são relevantes para a descontinuação do recurso em
    Microsoft Entra ID.

    O relatório é salvo em CSV e pode ser usado para avaliar impacto antes da
    mudança de suporte planejada pela Microsoft.

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1

    Executa a consulta usando o caminho padrão de saída: C:\M365-Report\Entra ID

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1 -OutputPath "D:\Reports\Entra"

    Executa a consulta e salva o CSV em um diretório personalizado, caso contrário será em "C:\M365-Report\Entra ID".

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1 -TenantId "<tenant-id>"

    Conecta a um tenant específico.

.NOTES
    Autor: Daniel Bessa
    Versão: 1.0
    Requisitos: Microsoft.Graph PowerShell module
    Permissões: Group.Read.All
#>

param(
    [string]$OutputPath = "C:\M365-Report\Entra ID",
    [string]$TenantId = ""
)

$ErrorActionPreference = "Stop"

function Write-Info($Message) {
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Warn($Message) {
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

try {
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell SDK não está instalado. Execute: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $existingContext) {
        if ($TenantId) {
            Write-Info "Conectando ao Microsoft Graph para o tenant $TenantId..."
            Connect-MgGraph -TenantId $TenantId -Scopes "Group.Read.All" -NoWelcome
        }
        else {
            Write-Info "Conectando ao Microsoft Graph..."
            Connect-MgGraph -Scopes "Group.Read.All" -NoWelcome
        }
    }
    else {
        Write-Info "Já existe uma sessão ativa do Microsoft Graph. Utilizando o contexto atual."
    }

    $targetFolder = [System.IO.Path]::GetFullPath($OutputPath)
    if (-not (Test-Path -Path $targetFolder)) {
        Write-Info "Criando pasta de saída: $targetFolder"
        New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
    }

    Write-Info "Consultando grupos dinâmicos..."

    $groups = Get-MgGroup -All -PageSize 999 `
        -Property Id, DisplayName, Description, GroupTypes, MembershipRule, MembershipRuleProcessingState, SecurityEnabled, MailEnabled `
        -Filter "groupTypes/any(c:c eq 'DynamicMembership')" `
        -ErrorAction Stop

    $results = foreach ($group in $groups) {
        $rule = [string]($group.MembershipRule)
        $usesMemberOf = $rule -match '(?i)(user\.)?memberOf|memberof'

        if ($usesMemberOf) {
            [pscustomobject]@{
                Id                            = $group.Id
                DisplayName                   = $group.DisplayName
                Description                   = $group.Description
                MembershipRule                = $rule
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                SecurityEnabled               = $group.SecurityEnabled
                MailEnabled                   = $group.MailEnabled
                GroupTypes                    = ($group.GroupTypes -join '; ')
                UsesMemberOf                  = $true
            }
        }
    }

    if (-not $results) {
        Write-Warn "Nenhum grupo dinâmico com regra de memberOf foi encontrado neste tenant."
        $results = @()
    }
    else {
        Write-Info ("Foram encontrados {0} grupos dinâmicos que usam memberOf." -f $results.Count)
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $targetFolder ("DynamicGroups-Using-MemberOf_{0}.csv" -f $timestamp)

    $results | Sort-Object DisplayName | Export-Csv -Path $reportPath -NoTypeInformation

    Write-Host "" 
    Write-Host "Resumo dos grupos dinâmicos que usam memberOf:" -ForegroundColor Green
    $results | Select-Object DisplayName, Id, MembershipRuleProcessingState, UsesMemberOf | Format-Table -AutoSize

    Write-Host "" 
    Write-Host "Arquivo CSV gerado em: $reportPath" -ForegroundColor Green

    return $results
}
catch {
    Write-Error "Falha ao consultar grupos dinâmicos: $($_.Exception.Message)"
    throw
}
