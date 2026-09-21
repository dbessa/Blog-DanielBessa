<#
.SYNOPSIS
    Identifica objetos do Microsoft Entra ID que usam regras memberOf ou políticas automáticas relevantes para a mudança de suporte do recurso.

.DESCRIPTION
    Este script conecta ao Microsoft Graph e consulta:
      - grupos dinâmicos
      - unidades administrativas dinâmicas
      - políticas de atribuição do Entitlement Management

    Em seguida, localiza regras/definições que utilizam memberOf, user.memberOf,
    ou configurações de atribuição automática, permitindo avaliar o impacto antes
    de mudanças de suporte no Microsoft Entra ID.

    O relatório é salvo em CSV e pode ser usado para revisão de impacto em tenant.

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1

    Executa a consulta usando o caminho padrão de saída: C:\M365-Report\Entra ID

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1 -OutputPath "D:\Reports\Entra"

    Executa a consulta e salva o CSV em um diretório personalizado.

.EXAMPLE
    .\EntraID-MemberOf-Check.ps1 -TenantId "<tenant-id>"

    Conecta a um tenant específico.

.NOTES
    Autor: Daniel Bessa
    Versão: 2.0
    Data de criação: 15/09/2026
    Atualização: 20/09/2026
    Requisitos: Microsoft.Graph PowerShell module
    Permissões recomendadas: Group.Read.All, Directory.Read.All, EntitlementManagement.Read.All
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

function ConvertTo-FlatString {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [string]) {
        return $Value
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [System.String])) {
        return ($Value | ForEach-Object { $_ }) -join '; '
    }

    return [string]$Value
}

function Test-MemberOfRule {
    param(
        [AllowEmptyString()][string]$Rule
    )

    if ([string]::IsNullOrWhiteSpace($Rule)) {
        return $false
    }

    return $Rule -match '(?i)(user\.)?memberOf|memberof'
}

function Test-AutomaticAssignmentPolicy {
    param(
        $Policy
    )

    if ($null -eq $Policy) {
        return $false
    }

    $serializableValues = @()
    foreach ($propertyName in @('AssignmentType', 'AllowedTargetScope', 'RequestorSettings', 'ApprovalSettings', 'Expiration', 'CanExtend', 'Question', 'IsDefault')) {
        if ($Policy.PSObject.Properties.Name -contains $propertyName) {
            $serializableValues += ConvertTo-FlatString -Value $Policy.$propertyName
        }
    }

    $combined = ($serializableValues -join ' ')
    return $combined -match '(?i)automatic|auto.*(assign|approve|accept|grant)'
}

try {
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw "Microsoft Graph PowerShell SDK não está instalado. Execute: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $existingContext) {
        if ($TenantId) {
            Write-Info "Conectando ao Microsoft Graph para o tenant $TenantId..."
            Connect-MgGraph -TenantId $TenantId -Scopes "Group.Read.All","Directory.Read.All","EntitlementManagement.Read.All" -NoWelcome
        }
        else {
            Write-Info "Conectando ao Microsoft Graph..."
            Connect-MgGraph -Scopes "Group.Read.All","Directory.Read.All","EntitlementManagement.Read.All" -NoWelcome
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

    $results = @()

    Write-Info "Consultando grupos dinâmicos..."
    $groups = Get-MgGroup -All -PageSize 999 `
        -Property Id, DisplayName, Description, GroupTypes, MembershipRule, MembershipRuleProcessingState, SecurityEnabled, MailEnabled `
        -Filter "groupTypes/any(c:c eq 'DynamicMembership')" `
        -ErrorAction Stop

    foreach ($group in $groups) {
        $rule = [string]($group.MembershipRule)
        if (Test-MemberOfRule -Rule $rule) {
            $results += [pscustomobject]@{
                SourceType                    = 'DynamicGroup'
                ObjectType                    = 'Group'
                Id                            = $group.Id
                DisplayName                   = $group.DisplayName
                Description                   = $group.Description
                MembershipRule                = $rule
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                SecurityEnabled               = $group.SecurityEnabled
                MailEnabled                   = $group.MailEnabled
                GroupTypes                    = ($group.GroupTypes -join '; ')
                UsesMemberOf                  = $true
                AssignmentType                = $null
                RequestorSettings             = $null
            }
        }
    }

    Write-Info "Consultando unidades administrativas dinâmicas..."
    $administrativeUnits = Get-MgDirectoryAdministrativeUnit -All -PageSize 999 `
        -Property Id, DisplayName, Description, MembershipRule, MembershipRuleProcessingState, MembershipType `
        -ErrorAction Stop

    foreach ($administrativeUnit in $administrativeUnits) {
        $rule = [string]($administrativeUnit.MembershipRule)
        if (Test-MemberOfRule -Rule $rule) {
            $results += [pscustomobject]@{
                SourceType                    = 'DynamicAdministrativeUnit'
                ObjectType                    = 'AdministrativeUnit'
                Id                            = $administrativeUnit.Id
                DisplayName                   = $administrativeUnit.DisplayName
                Description                   = $administrativeUnit.Description
                MembershipRule                = $rule
                MembershipRuleProcessingState = $administrativeUnit.MembershipRuleProcessingState
                SecurityEnabled               = $null
                MailEnabled                   = $null
                GroupTypes                    = ConvertTo-FlatString -Value $administrativeUnit.MembershipType
                UsesMemberOf                  = $true
                AssignmentType                = $null
                RequestorSettings             = $null
            }
        }
    }

    Write-Info "Consultando políticas de atribuição do Entitlement Management..."
    $assignmentPolicyCommand = Get-Command Get-MgEntitlementManagementAssignmentPolicy -ErrorAction SilentlyContinue
    if ($assignmentPolicyCommand) {
        $assignmentPolicies = Get-MgEntitlementManagementAssignmentPolicy -All -PageSize 999 `
            -Property Id, DisplayName, Description, AssignmentType, AllowedTargetScope, RequestorSettings, ApprovalSettings, Expiration, IsDefault `
            -ErrorAction Stop

        foreach ($policy in $assignmentPolicies) {
            $isAutomaticAssignment = Test-AutomaticAssignmentPolicy -Policy $policy
            if ($isAutomaticAssignment) {
                $results += [pscustomobject]@{
                    SourceType                    = 'EntitlementManagementAssignmentPolicy'
                    ObjectType                    = 'EntitlementManagement'
                    Id                            = $policy.Id
                    DisplayName                   = $policy.DisplayName
                    Description                   = $policy.Description
                    MembershipRule                = ConvertTo-FlatString -Value $policy.AssignmentType
                    MembershipRuleProcessingState = $null
                    SecurityEnabled               = $null
                    MailEnabled                   = $null
                    GroupTypes                    = ConvertTo-FlatString -Value $policy.AssignmentType
                    UsesMemberOf                  = $false
                    AssignmentType                = ConvertTo-FlatString -Value $policy.AssignmentType
                    RequestorSettings             = ConvertTo-FlatString -Value $policy.RequestorSettings
                }
            }
        }
    }
    else {
        Write-Warn "O cmdlet Get-MgEntitlementManagementAssignmentPolicy não está disponível neste módulo do Microsoft.Graph. Instale a versão mais recente do módulo para incluir essa verificação."
    }

    if (-not $results) {
        Write-Warn "Nenhum objeto com regra memberOf ou política automática relevante foi encontrado neste tenant."
        $results = @()
    }
    else {
        Write-Info ("Foram encontrados {0} objetos relevantes para revisão." -f $results.Count)
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $targetFolder ("EntraID-Impact-Review_{0}.csv" -f $timestamp)

    $results | Sort-Object SourceType, DisplayName | Export-Csv -Path $reportPath -NoTypeInformation

    Write-Host "" 
    Write-Host "Resumo dos objetos relevantes:" -ForegroundColor Green
    $results | Select-Object SourceType, DisplayName, Id, AssignmentType, MembershipRuleProcessingState | Format-Table -AutoSize

    Write-Host "" 
    Write-Host "Arquivo CSV gerado em: $reportPath" -ForegroundColor Green

    return $results
}
catch {
    Write-Error "Falha ao consultar objetos do Microsoft Entra ID: $($_.Exception.Message)"
    throw
}
