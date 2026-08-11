<#
================================================================================
 Entra App / Graph API 인벤토리 수집 스크립트
--------------------------------------------------------------------------------
 M365 관리센터의 "앱 등록 내보내기" CSV에는 아래 항목이 빠져 있어
 깡통 App 판정과 Graph API 미사용/오용 검토가 불가능하다.
     - requiredResourceAccess  (요청한 API 권한 목록)   ★핵심
     - owners                  (소유자)                 ★메일 발송 대상
     - appRoleAssignments      (실제 동의된 앱 권한)
     - oauth2PermissionGrants  (실제 동의된 위임 권한)
     - signInActivity          (마지막 사용 시각)

 이 스크립트는 위 항목을 Microsoft Graph에서 직접 수집해 CSV 5종으로 저장한다.

 실행 방법
     1) 최초 1회만
        Install-Module Microsoft.Graph      -Scope CurrentUser
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser   # 로그인 활동용(선택)
     2) 실행
        .\collect_entra_inventory.ps1 -OutDir .\entra_export

 필요 권한 (읽기 전용)
     Application.Read.All / Directory.Read.All / AuditLog.Read.All
     ※ Connect-MgGraph 대화형 로그인은 실행자 본인 권한으로 동작하는 Delegated 방식이라
       별도 Entra App 등록·앱 권한 부여가 불필요하다.
       사내 등급 기준으로는 접근범위 D → G1에 해당한다.
================================================================================
#>
[CmdletBinding()]
param(
    [string]$OutDir = ".\entra_export",
    [switch]$SkipSignInActivity   # AuditLog 권한이 없거나 P1 미보유 시
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force $OutDir | Out-Null }
$stamp = Get-Date -Format "yyyyMMdd"
function Out-Csv($obj, $name) {
    $p = Join-Path $OutDir "$($name)_$stamp.csv"
    $obj | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8
    Write-Host ("  저장: {0,-34} {1,5}행" -f (Split-Path $p -Leaf), @($obj).Count) -ForegroundColor Green
}

Write-Host "`n[0] Microsoft Graph 연결" -ForegroundColor Cyan
Connect-MgGraph -Scopes "Application.Read.All","Directory.Read.All","AuditLog.Read.All" -NoWelcome
$ctx = Get-MgContext
Write-Host "  테넌트: $($ctx.TenantId) / 계정: $($ctx.Account)"

# ── Graph 권한 GUID → 이름 사전 ───────────────────────────────────────────────
# requiredResourceAccess 에는 권한 이름이 아니라 GUID가 들어 있어 반드시 변환이 필요하다.
Write-Host "`n[1] 리소스 API 권한 사전 구성" -ForegroundColor Cyan
$permMap = @{}   # "<resourceAppId>|<permId>" -> @{Name; Type; Resource}
$resourceNames = @{}
$knownResources = @(
    "00000003-0000-0000-c000-000000000000",  # Microsoft Graph
    "00000002-0000-0000-c000-000000000000",  # Azure AD Graph (레거시)
    "00000003-0000-0ff1-ce00-000000000000",  # SharePoint
    "00000002-0000-0ff1-ce00-000000000000",  # Exchange Online
    "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe",  # Teams Services
    "00000009-0000-0000-c000-000000000000"   # Power BI Service
)
foreach ($rid in $knownResources) {
    try {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$rid'" -Property Id,AppId,DisplayName,AppRoles,Oauth2PermissionScopes -ErrorAction Stop
        if (-not $sp) { continue }
        $resourceNames[$rid] = $sp.DisplayName
        foreach ($r in $sp.AppRoles)              { $permMap["$rid|$($r.Id)"] = @{ Name=$r.Value; Type="Application"; Resource=$sp.DisplayName } }
        foreach ($s in $sp.Oauth2PermissionScopes) { $permMap["$rid|$($s.Id)"] = @{ Name=$s.Value; Type="Delegated";  Resource=$sp.DisplayName } }
    } catch { Write-Warning "  리소스 $rid 조회 실패: $($_.Exception.Message)" }
}
Write-Host "  권한 사전: $($permMap.Count)건 / 리소스 $($resourceNames.Count)종"

# ── 앱 등록 + 요청 권한 + 소유자 ─────────────────────────────────────────────
Write-Host "`n[2] 앱 등록 · 요청 권한 · 소유자 수집" -ForegroundColor Cyan
$apps = Get-MgApplication -All -Property Id,AppId,DisplayName,CreatedDateTime,SignInAudience,Tags,`
    RequiredResourceAccess,PasswordCredentials,KeyCredentials,Web,PublicClient,Notes
Write-Host "  앱 등록: $($apps.Count)건"

$appRows  = New-Object System.Collections.Generic.List[object]
$permRows = New-Object System.Collections.Generic.List[object]
$ownerRows= New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($a in $apps) {
    $i++
    if ($i % 50 -eq 0) { Write-Host "    ...$i / $($apps.Count)" }

    # 소유자
    $owners = @()
    try {
        $owners = Get-MgApplicationOwner -ApplicationId $a.Id -All -ErrorAction Stop
    } catch { }
    foreach ($o in $owners) {
        $ap = $o.AdditionalProperties
        $ownerRows.Add([pscustomobject]@{
            appId=$a.AppId; appDisplayName=$a.DisplayName; ownerObjectId=$o.Id
            ownerType=($ap["@odata.type"] -replace "#microsoft.graph.","")
            ownerDisplayName=$ap["displayName"]; ownerUPN=$ap["userPrincipalName"]; ownerMail=$ap["mail"]
        })
    }

    # 요청 권한 (requiredResourceAccess)
    $permCount = 0; $graphPermCount = 0; $names = @()
    foreach ($rra in $a.RequiredResourceAccess) {
        foreach ($ra in $rra.ResourceAccess) {
            $permCount++
            $k = "$($rra.ResourceAppId)|$($ra.Id)"
            $m = $permMap[$k]
            $pname = if ($m) { $m.Name } else { "(미해석:$($ra.Id))" }
            $ptype = if ($ra.Type -eq "Role") { "Application" } else { "Delegated" }
            $res   = if ($resourceNames.ContainsKey($rra.ResourceAppId)) { $resourceNames[$rra.ResourceAppId] } else { $rra.ResourceAppId }
            if ($rra.ResourceAppId -eq "00000003-0000-0000-c000-000000000000") { $graphPermCount++; $names += $pname }
            $permRows.Add([pscustomobject]@{
                appId=$a.AppId; appDisplayName=$a.DisplayName
                resourceAppId=$rra.ResourceAppId; resource=$res
                permissionId=$ra.Id; permissionName=$pname; permissionType=$ptype
            })
        }
    }

    $creds = @($a.PasswordCredentials) + @($a.KeyCredentials)
    $ends  = $creds | Where-Object { $_.EndDateTime } | ForEach-Object { $_.EndDateTime }
    $maxEnd= if ($ends) { ($ends | Measure-Object -Maximum).Maximum } else { $null }
    $redir = @($a.Web.RedirectUris) + @($a.PublicClient.RedirectUris)

    $appRows.Add([pscustomobject]@{
        appId=$a.AppId; objectId=$a.Id; displayName=$a.DisplayName
        createdDateTime=$a.CreatedDateTime; ageDays=[int]((Get-Date) - $a.CreatedDateTime).TotalDays
        signInAudience=$a.SignInAudience; tags=($a.Tags -join ";")
        credentialCount=$creds.Count; credentialMaxExpiry=$maxEnd
        credentialState= if ($creds.Count -eq 0) { "없음" }
                         elseif ($maxEnd -and $maxEnd -lt (Get-Date)) { "만료" }
                         elseif ($maxEnd -and $maxEnd -lt (Get-Date).AddDays(90)) { "90일내만료" }
                         else { "유효" }
        redirectUriCount=$redir.Count
        totalPermissionCount=$permCount
        graphPermissionCount=$graphPermCount
        graphPermissions=($names | Sort-Object -Unique) -join ";"
        ownerCount=$owners.Count
        ownerMails=(($owners | ForEach-Object { $_.AdditionalProperties["mail"] } | Where-Object { $_ }) -join ";")
        notes=$a.Notes
    })
}
Out-Csv $appRows   "01_applications"
Out-Csv $permRows  "02_requested_permissions"
Out-Csv $ownerRows "03_owners"

# ── 서비스 주체 + 실제 동의된 권한 ───────────────────────────────────────────
Write-Host "`n[3] 서비스 주체 · 실제 동의된 권한 수집" -ForegroundColor Cyan
$sps = Get-MgServicePrincipal -All -Property Id,AppId,DisplayName,ServicePrincipalType,AccountEnabled,Tags,AppOwnerOrganizationId
$myTenant = $ctx.TenantId
$grantRows = New-Object System.Collections.Generic.List[object]

# 앱 권한 (Application permissions) — 관리자 동의로 부여된 것
$i = 0
foreach ($sp in $sps) {
    $i++
    if ($i % 100 -eq 0) { Write-Host "    ...$i / $($sps.Count)" }
    try {
        $ras = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction Stop
    } catch { continue }
    foreach ($ra in $ras) {
        $k = $null
        $resSp = $sps | Where-Object { $_.Id -eq $ra.ResourceId } | Select-Object -First 1
        $resAppId = if ($resSp) { $resSp.AppId } else { $null }
        $pname = if ($resAppId -and $permMap.ContainsKey("$resAppId|$($ra.AppRoleId)")) { $permMap["$resAppId|$($ra.AppRoleId)"].Name } else { "(미해석)" }
        $grantRows.Add([pscustomobject]@{
            appId=$sp.AppId; spObjectId=$sp.Id; appDisplayName=$sp.DisplayName
            grantType="Application"; resource=$ra.ResourceDisplayName
            permissionName=$pname; permissionId=$ra.AppRoleId
            principal=""; createdDateTime=$ra.CreatedDateTime
        })
    }
}
# 위임 권한 (Delegated) — 사용자/관리자 동의
foreach ($g in (Get-MgOauth2PermissionGrant -All)) {
    $cli = $sps | Where-Object { $_.Id -eq $g.ClientId } | Select-Object -First 1
    $res = $sps | Where-Object { $_.Id -eq $g.ResourceId } | Select-Object -First 1
    foreach ($scope in ($g.Scope -split "\s+" | Where-Object { $_ })) {
        $grantRows.Add([pscustomobject]@{
            appId=$cli.AppId; spObjectId=$g.ClientId; appDisplayName=$cli.DisplayName
            grantType="Delegated"; resource=$res.DisplayName
            permissionName=$scope; permissionId=""
            principal=$(if ($g.ConsentType -eq "AllPrincipals") { "전체(관리자동의)" } else { $g.PrincipalId })
            createdDateTime=""
        })
    }
}
Out-Csv $grantRows "04_granted_permissions"

# ── 서비스 주체 로그인 활동 (= 실제 사용 여부) ───────────────────────────────
if (-not $SkipSignInActivity) {
    Write-Host "`n[4] 서비스 주체 로그인 활동 수집" -ForegroundColor Cyan
    Write-Host "    ※ Microsoft Entra ID P1 이상 필요. 실패하면 -SkipSignInActivity 로 재실행하세요." -ForegroundColor DarkYellow
    try {
        $act = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?`$top=999"
        $rows = New-Object System.Collections.Generic.List[object]
        do {
            foreach ($x in $act.value) {
                $rows.Add([pscustomobject]@{
                    appId=$x.appId
                    lastSignIn                = $x.lastSignInActivity.lastSignInDateTime
                    lastAppOnlySignIn         = $x.applicationAuthenticationClientSignInActivity.lastSignInDateTime
                    lastDelegatedClientSignIn = $x.delegatedClientSignInActivity.lastSignInDateTime
                    lastDelegatedResourceSignIn=$x.delegatedResourceSignInActivity.lastSignInDateTime
                })
            }
            $next = $act.'@odata.nextLink'
            if ($next) { $act = Invoke-MgGraphRequest -Method GET -Uri $next }
        } while ($next)
        Out-Csv $rows "05_signin_activity"
    } catch {
        Write-Warning "  로그인 활동 수집 실패: $($_.Exception.Message)"
        Write-Host "  → Entra 관리센터 > 엔터프라이즈 애플리케이션 > 사용량 및 인사이트 에서 수동 내보내기로 대체하세요." -ForegroundColor DarkYellow
    }
}

# ── 요약 ─────────────────────────────────────────────────────────────────────
Write-Host "`n────────────────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host " 수집 요약" -ForegroundColor Cyan
Write-Host "────────────────────────────────────────────────────────────" -ForegroundColor Cyan
$noGraph  = @($appRows | Where-Object { $_.graphPermissionCount -eq 0 }).Count
$noCred   = @($appRows | Where-Object { $_.credentialState -eq "없음" }).Count
$noOwner  = @($appRows | Where-Object { $_.ownerCount -eq 0 }).Count
$hollow   = @($appRows | Where-Object { $_.graphPermissionCount -eq 0 -and $_.credentialState -eq "없음" -and $_.redirectUriCount -eq 0 }).Count
Write-Host ("  전체 앱 등록                        {0,5}" -f $appRows.Count)
Write-Host ("  Graph API 권한 없음                 {0,5}" -f $noGraph)
Write-Host ("  자격증명 없음 (토큰 발급 불가)      {0,5}" -f $noCred)
Write-Host ("  소유자 없음 (연락 대상 불명)        {0,5}" -f $noOwner) -ForegroundColor Yellow
Write-Host ("  깡통 후보 (권한·자격증명·URI 전무)  {0,5}" -f $hollow) -ForegroundColor Red
Write-Host ("  요청 권한 총건수                    {0,5}" -f $permRows.Count)
Write-Host ("  실제 동의된 권한 건수               {0,5}" -f $grantRows.Count)
Write-Host "`n  출력 폴더: $((Resolve-Path $OutDir).Path)"
Write-Host "  → 이 CSV들을 GAPS Tab3(회수 대상 산출기)에 업로드하면 회수 대상과 소유자별 메일 초안이 생성됩니다.`n"
Disconnect-MgGraph | Out-Null
