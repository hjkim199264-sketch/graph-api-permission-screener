# PRD — Graph API 권한 등급 자동 판정 시스템 (GAPS v1)

| 항목 | 내용 |
|---|---|
| 제품명 | **GAPS** — Graph API Permission Screener |
| 버전 | v1.0 |
| 작성일 | 2026-08-11 |
| 상위 문서 | [BRD_GraphAPI_권한관리체계.md](BRD_GraphAPI_권한관리체계.md) |
| 산출물 | `GraphAPI_등급판정시스템.html` (단일 파일, 오프라인) |
| 주 사용자 | DX운영그룹(M365 운영) 담당자 — 부 사용자: 요청자 셀프체크, 정보보호그룹 |

---

## 1. 제품 개요

### 1-1. One-liner

> **신청 템플릿 PPTX를 끌어다 놓으면, API별 권한 등급(G1/G2/G3) 초안 · 보완 요청사항 · 다음 진행 절차를 즉시 산출하는 오프라인 단일 페이지 도구.**

### 1-2. 처리 흐름

```
[PPTX 업로드]
      │  브라우저 내에서만 처리 (외부 전송 없음)
      ▼
[1] ZIP 해제 → ppt/slides/slideN.xml 추출
      ▼
[2] 표(a:tbl)·텍스트박스 파싱 → 라벨 기반 필드 매핑
      ▼
[3] API 항목 분해 (①②③… 복수 API 지원)
      ▼
[4] 1차 분류 : 정보유형(G/S) × 접근범위(D/A) × 작업유형(R/W/X) → 8조합 → G1/G2/G3
      │      └─ ★ 접근범위 재판정 룰 (Application → Delegated)
      ▼
[5] 2차 분류 : 4개 판단 지표로 상·하향 → 최종 등급
      ▼
[6] 신청 건 종합 등급 = MAX(API별 등급)
      ▼
[7] 보완 요청사항 생성 (누락 / 최소권한 / 고위험 / 운영위생)
      ▼
[8] 향후 진행사항 안내 (등급별 절차·담당·소요)
      ▼
[결과 화면 + IR 회신문 복사 + JSON 내보내기]
```

### 1-3. 설계 원칙

| # | 원칙 | 구현 |
|---|---|---|
| D1 | **결정적(deterministic)** | 규칙 기반. AI/확률 추론 없음 → 동일 입력 = 동일 출력, 감사 가능 |
| D2 | **근거 제시** | 모든 판정에 발동한 규칙 ID와 원문 근거 문장을 병기 |
| D3 | **초안일 뿐** | 화면·출력물 전면에 Draft 배지. 자동 승인 경로 없음 |
| D4 | **보수적 기본값** | 미확인 API·모호 항목은 **높은 등급 쪽**으로 가정하고 확인 요청 |
| D5 | **오프라인/무전송** | 외부 CDN·네트워크 호출 0건. `DecompressionStream` 네이티브 + 자체 inflate 폴백 |
| D6 | **수기 보정 허용** | 파싱 결과를 화면에서 수정하면 즉시 재판정 |

---

## 2. 사용자 스토리

| ID | As a … | I want to … | So that … | 우선순위 |
|---|---|---|---|---|
| US-1 | DX운영 담당자 | 신청서를 올리는 즉시 등급 초안을 본다 | 접수 당일 1차 회신할 수 있다 | Must |
| US-2 | DX운영 담당자 | 미기재 항목을 목록으로 받는다 | 한 번에 모아서 보완 요청할 수 있다 | Must |
| US-3 | DX운영 담당자 | Application 과대 신청을 자동 탐지받는다 | 최소권한으로 협의를 시작할 수 있다 | Must |
| US-4 | DX운영 담당자 | 판정 결과를 IR 회신문으로 복사한다 | 재작성 없이 그대로 회신한다 | Should |
| US-5 | 요청자 | 제출 전에 스스로 돌려본다 | 반려 없이 한 번에 접수된다 | Should |
| US-6 | 정보보호그룹 | 표준 보안 체크리스트를 함께 받는다 | 검토 관점이 매번 동일하다 | Must |
| US-7 | 담당자 | 파싱이 틀리면 화면에서 고친다 | 템플릿 변형에도 도구를 쓸 수 있다 | Must |
| US-8 | 담당자 | 결과를 JSON으로 내려받는다 | 요청 대장·사후 감사 근거로 남긴다 | Should |

---

## 3. 기능 요구사항

### FR-1. 파일 입력

| ID | 요구사항 |
|---|---|
| FR-1.1 | `.pptx` 파일을 드래그&드롭 또는 파일 선택으로 입력 |
| FR-1.2 | 브라우저 내에서만 처리. 어떠한 네트워크 요청도 발생시키지 않음 |
| FR-1.3 | `.ppt`(레거시), 암호화/DRM(IRM) 적용 파일은 명확한 안내 메시지와 함께 거부 |
| FR-1.4 | 전 슬라이드를 스캔하되 신청 템플릿 표가 있는 슬라이드를 자동 식별 |
| FR-1.5 | **샘플 데이터로 즉시 체험** 버튼 제공 (파일 없이 동작 확인) |

**DRM 판별**: 파일 선두 4바이트가 `D0 CF 11 E0`(OLE Compound File)이면 암호화/IRM 적용으로 간주하고 안내한다.

### FR-2. 필드 추출 (파싱)

**라벨 기반 매칭** — 셀 위치가 아닌 라벨 문구로 찾는다. Full/약식 템플릿 및 행 순서 변경에 모두 대응.

| 정규화 필드 | 매칭 라벨(정규식, 공백·`*`·번호 제거 후) |
|---|---|
| `taskName` | `과제명`, `프로젝트명` |
| `taskOverview` | `과제개요`, `과제 개요`, `개요` |
| `requester` | `요청자/부서`, `요청자`, `신청자` |
| `developer` | `개발자/소속`, `개발자` |
| `api[n].nameAndPurpose` | `①API명/활용목적`, `①요청API명/활용목적`, `API명/활용목적` |
| `api[n].data` | `①활용데이터/URL`, `활용데이터/URL`, `활용데이터` |
| `api[n].appId` | `①EntraAPPID`, `EntraAPPID`, `AppID` |
| `api[n].permType` | `①Delegated/Application`, `Delegated/Application`, `권한유형` |
| `api[n].account` | `①활용계정/App`, `활용계정/App`, `활용계정` |
| `scopeVolume` | `영향범위/활용규모`, `영향범위`, `활용규모` |
| `consent` | `정보주체인지/동의`, `정보주체동의`, `동의` |
| `storage` | `저장/2차가공`, `저장2차가공`, `저장및2차가공` |
| `remark` | `비고` |
| `expiry` | `권한만료일`, `만료일`, `사용기한` |

**값 추출 규칙**

1. 라벨 셀 기준 **같은 행에서 오른쪽으로 스캔**, 첫 번째 비어있지 않은 셀을 값으로 채택 (병합 셀 대응).
2. 오른쪽에 값이 없으면 **바로 아래 행 같은 열**을 확인 (세로형 템플릿 대응).
3. 표 밖 텍스트박스에서 `라벨 : 값`, `라벨 - 값` 패턴도 보조 추출.
4. `T.B.D`, `TBD`, `미정`, `-`, `N/A`(단, `storage`/`consent`는 유효값), `?`, 빈 문자열 → **미기재**로 판정.

**복수 API 처리**: 라벨의 원문자(`①②③④⑤`) 또는 숫자 접두(`1)`, `2.`)로 인덱스를 인식하여 API 슬롯을 분리한다. 인덱스가 없으면 단일 API로 처리한다.

**API 권한명 추출**: `nameAndPurpose` 값에서 `[A-Za-z][A-Za-z0-9]*(\.[A-Za-z][A-Za-z0-9]*){1,3}` 패턴으로 권한 후보를 뽑고, 사전 매칭 + 접미사 검증(`.Read`, `.ReadWrite`, `.All`, `.Send` …)으로 확정한다. `:` 이후는 활용목적으로 분리한다.

### FR-3. 1차 분류

#### FR-3.1 정보 유형 (G / S)

내장 **API 사전**으로 판정. 사전은 슬라이드 7 「주요 Graph API 등급 분류 초안」 23건을 정본으로 하고, 동일 리소스 패밀리로 확장한다.

| infoType | 의미 | 근거 |
|---|---|---|
| `G` | 일반 정보 확정 | 슬라이드 7에서 `G1/G2` 또는 `G2`로 표기된 권한 |
| `S` | 민감 정보 확정 | 슬라이드 7에서 `G3`으로 표기된 권한 |
| `GS` | 활용 데이터에 따라 결정 | 슬라이드 7에서 `G2/G3`으로 표기된 권한 |

> **`G1/G2` 표기의 해석**: 접근범위가 D면 G1, A면 G2 → 정보유형은 `G` 확정.
> **`G2/G3` 표기의 해석**: Application 전제에서 정보유형이 활용 데이터에 따라 갈림 → `GS`.
> `GS`는 **보수적으로 `S`로 가정**하고 "활용 필드를 명시해 주십시오"를 보완 요청에 추가한다.

사전 미등재 권한 → `UNKNOWN` → `S` 가정 + 수동 지정 요청.

#### FR-3.2 접근 범위 (D / A) — ★ 핵심 룰

```
declared = Delegated  →  D
declared = Application →  RULE-SCOPE-D 평가
```

**RULE-SCOPE-D (Application → Delegated 재판정)**

- **평가 대상 텍스트** (해당 API 슬롯 한정):
  `api[n].nameAndPurpose` + `api[n].data` + `api[n].account`
  (보조: `taskOverview`, `remark` — 보조 근거는 단독으로 발동시키지 않고 가중치만 부여)

- **발동 키워드 (LIMIT)**
  `특정 사서함`, `특정 메일함`, `특정 팀`, `특정 채널`, `특정 사이트`, `특정 사용자`, `특정 계정`, `특정 부서`, `특정 그룹`, `특정 공유사서함`, `공유 사서함`, `공유사서함`, `단일 사서함`, `지정된 사서함`, `지정 사이트`, `개인 사서함`, `본인 사서함`, `본인 계정`, `개인 계정`, `로그인한 사용자`, `로그인된 사용자`, `로그인 된 사용자`, `해당 사용자`, `사용자 본인`, `담당자 1인`, `1인의`, `일부 사이트`, `해당 팀`, `해당 채널`, `Sites.Selected`

- **차단 키워드 (TENANT)** — 같은 텍스트에 있으면 재판정하지 않음
  `전사`, `전체 사용자`, `모든 사용자`, `전 사서함`, `전체 사서함`, `전 임직원`, `임직원 전체`, `테넌트 전체`, `조직 전체`, `all users`, `all mailboxes`, `전체 사이트`, `모든 팀`

- **판정**
  | LIMIT | TENANT | 결과 |
  |---|---|---|
  | O | X | **D로 재판정** + `RULE-SCOPE-D` 근거 문장 표시 + 보완요청 `SEC-SCOPE` **필수** 부착 |
  | O | O | **A 유지** + "대상 범위 상충 — 명확화 필요" 보완요청 |
  | X | — | A 유지 |

- **부착되는 필수 조건** (`SEC-SCOPE`)
  > 재판정으로 하향된 등급은 **아래 기술적 범위 제한 조치의 이행을 전제**로 한다. 미이행 시 원 등급(재판정 전)으로 환원한다.

  | 워크로드 | 조치 |
  |---|---|
  | Exchange Online | **RBAC for Applications**(권장) 또는 `New-ApplicationAccessPolicy`로 대상 사서함(보안그룹) 한정 |
  | SharePoint/OneDrive | `Sites.Selected` 권한 + 대상 사이트에만 권한 부여 |
  | Teams | **RSC**(Resource-Specific Consent) / `*.Group` 계열 권한으로 특정 팀에만 동의 |
  | 공통 | 사용자 로그인 컨텍스트가 존재하면 **Delegated 권한으로 전환**(최우선 대안) |

  ※ 근거: Application 권한은 별도 제한이 없으면 **테넌트 전체**에 적용된다. "특정 사서함만 쓰겠다"는 선언은 애플리케이션 코드 레벨의 약속일 뿐, **부여된 권한 자체를 좁히지 않는다.**

#### FR-3.3 작업 유형 (R / W / X)

| 판정 | 조건 (권한명 기준, 우선순위 순) |
|---|---|
| **X** | 권한명에 `Delete`/`Purge` 포함, **또는** 활용목적에 `삭제`·`제거`·`퇴직`·`회수`·`파기`·`delete`·`purge` 포함 |
| **W** | `ReadWrite`, `.Write`, `.Send`, `.Create`, `.Manage`, `FullControl`, `EnableDisableAccount`, `ManageIdentities`, `.Invite` 포함 |
| **R** | 그 외 (`.Read`, `.ReadBasic`, `.Selected` 등) |

- `X` 발생 시 등급은 변경하지 않고 **`별도 협의` 플래그**를 세운다 (슬라이드 4 각주 준거).
- `ReadWrite` 계열은 삭제 권한을 내포하므로 `NOTE-RW-DELETE` 안내를 함께 출력한다.

#### FR-3.4 8조합 매핑

| 조합 | 등급 | | 조합 | 등급 |
|---|---|---|---|---|
| `GDR` | **G1** | | `GAR` | **G2** |
| `GDW` | **G1** | | `GAW` | **G2** |
| `SDR` | **G1** | | `SAR` | **G3** |
| `SDW` | **G1** | | `SAW` | **G3** |

구현: `범위 == D → G1` / `범위 == A && 정보 == G → G2` / `범위 == A && 정보 == S → G3`

### FR-4. 2차 분류 (최종 등급)

| 규칙 ID | 지표 | 발동 조건 | 조정 |
|---|---|---|---|
| `ADJ-MIN` | 데이터 최소화 | `api.data` 또는 `remark`에 `필터링`, `마스킹`, `제외`, `최소화`, `필수 필드만`, `일부 필드`, `본문 제외`, `메타데이터만`, `헤더만`, `ID만` | G3 → G2 |
| `ADJ-CONSENT` | 정보주체 인지/동의 | `consent`가 `전원 동의`, `동의`, `인지`, `공지 완료`, `Y`, `O` 등 긍정값 | G2 → G1 |
| `ADJ-STORE` | 저장·2차 가공 | `storage`가 `N/A`, `없음`, `미저장`, `저장 안 함`, `통계만`, `집계만` 등 **비저장** | G3 → G2 |
| `ADJ-SCOPE-UP` | 영향 범위/규모 | `scopeVolume`에 **(중요계정)** `임원`·`경영진`·`직책자`·`대표이사`·`CxO` **또는** **(전사범위 ∧ 고빈도)** (`전사`\|`전체 사용자`\|`전 임직원`) ∧ (`Daily`\|`매일`\|`일별`\|`실시간`\|`배치`\|`시간 단위`) | G2 → G3 |

**적용 정책 (설계 결정)**

1. 하향은 **최대 1등급**으로 제한한다. `ADJ-MIN`과 `ADJ-STORE`가 동시에 발동해도 G3 → G2까지만, 여기에 `ADJ-CONSENT`가 더해져도 G1로 내려가지 않는다. 이때 `ADJ-CAP` 안내를 출력한다.
   *근거: 2단계 하향은 "전사 민감"을 "개인 권한"과 동일 취급하는 것이어서 위험. 담당자 재량으로만 허용.*
2. 조정 순서는 **하향 먼저, 상향 나중**이다. 따라서 G3 → (`ADJ-MIN`) → G2 → (`ADJ-SCOPE-UP`) → G3 처럼 되돌아올 수 있으며, 이는 의도된 동작이다.
3. `ADJ-SCOPE-UP`은 **G2 → G3 자동 상향만** 수행한다. 조정 시점 등급이 G1인 경우에는 자동 상향하지 않고 `담당자 확인 권고` 안내만 출력한다.
   *근거: 원본 기준(슬라이드 4)이 "G2 → G3 로 상향 검토"만 규정하고 있고, G1은 Delegated(사용자 본인 권한 범위)라 규모 위험이 구조적으로 제한되기 때문.*
4. 상향과 하향이 동시 발동하면 **상쇄**(1차 등급 유지)하고 `ADJ-OFFSET` — `담당자 확인 필요`를 표시한다.
5. 모든 조정은 **발동 규칙 ID + 근거 원문**을 함께 출력한다.
6. `consent`/`storage`가 **미기재**면 조정하지 않고 보완 요청으로 넘긴다. (미기재 ≠ 유리한 해석)
   단, 이 두 항목에서 **`N/A`·`없음`은 미기재가 아니라 의미 있는 값**이다.
   - `storage: N/A` → **비저장**으로 해석 → `ADJ-STORE` 하향 근거 + `NOTE-NOSTORE`(이행 확인 대상) 출력
   - `consent: N/A`·`미정` → **동의 미확보**로 해석 → `REQ-CONSENT-NO` 출력. 심각도는 최종 등급 연동 — **G2 이상이면 `Blocker`, G1이면 `High`**

### FR-5. 신청 건 종합 등급

```
overallGrade = MAX(api[i].finalGrade)        // G3 > G2 > G1
```
- 하나라도 `별도 협의(X)` 플래그가 있으면 신청 건 전체에 표시.
- API별 등급과 종합 등급을 모두 노출한다.

### FR-6. 보완 요청사항 생성

심각도 4단계: `Blocker`(접수 보류) / `Critical` / `High` / `Medium`

#### A. 필수 항목 누락 — `Blocker`

| 규칙 ID | 항목 | 메시지 |
|---|---|---|
| `REQ-TASK` | 과제명·과제개요 | 과제 배경과 목적을 기재해 주십시오. |
| `REQ-API` | API명 | Microsoft Graph 공식 권한명으로 기재해 주십시오. (예: `Mail.Read`) |
| `REQ-PURPOSE` | 활용목적 | 해당 API로 **무엇을 하는지** 한 문장으로 기재해 주십시오. |
| `REQ-PERMTYPE` | Delegated/Application | 권한 유형을 명시해 주십시오. 미기재 시 Application(보수적)으로 가정합니다. |
| `REQ-DATA` | 활용 데이터/URL | 실제로 읽거나 쓰는 **데이터 필드**를 명시해 주십시오. |
| `REQ-APPID` | Entra App ID | Entra 앱 등록 후 App(Client) ID를 기재해 주십시오. |
| `REQ-SCOPE` | 영향 범위/활용 규모 | 영향 계정 수·중요도, 호출 주기를 기재해 주십시오. |
| `REQ-CONSENT` | 정보주체 인지/동의 | 데이터 주체의 인지·동의 확보 여부와 방법을 기재해 주십시오. (미기재 시 `Blocker`, 기재값이 판정 불가 시 `High`) |
| `REQ-CONSENT-NO` | 정보주체 인지/동의 | 동의 미확보(`N/A`·`미정`·`없음`)로 기재됨. **G2 이상 `Blocker` / G1 `High`** |
| `REQ-STORE` | 저장/2차 가공 | 저장 여부, 저장 시 위치·보관기간·2차 활용 범위를 기재해 주십시오. |
| `REQ-EXPIRY` | 권한 만료일 | **필수**. 미기재 신청은 접수 보류합니다. |
| `REQ-OWNER` | 요청자/부서, 개발자/소속 | 외부 개발사 인력 포함 시 소속을 명시해 주십시오. |

#### B. 최소권한 — `High`

| 규칙 ID | 조건 | 권고 |
|---|---|---|
| `SEC-SCOPE` | RULE-SCOPE-D 발동 | 기술적 범위 제한 조치 필수 (FR-3.2 표) |
| `SEC-ALL` | Application ∧ 권한명에 `.All` | 테넌트 전체 적용. 범위 한정 권한 또는 정책 기반 제한 검토 |
| `SEC-OVERPERM` | 권한 `W` ∧ 활용목적이 조회성(`조회`,`읽기`,`수집`,`대시보드`,`분석`,`표시`)만 | Read 전용 권한으로 축소 권고 (`Mail.ReadWrite` → `Mail.Read`) |
| `SEC-UNDERPERM` | 권한 `R` ∧ 활용목적에 변경성(`발송`,`생성`,`수정`,`업로드`,`할당`,`초대`) | 권한 부족 가능성. 목적/권한 정합성 재확인 |
| `SEC-MAILSEND` | `Mail.ReadWrite` ∧ 목적이 발송만 | `Mail.Send`로 대체 권고 (읽기 권한 불필요) |
| `SEC-SITES` | `Sites.Read.All` / `Sites.ReadWrite.All` | `Sites.Selected` + 대상 사이트 지정 권고 |
| `SEC-FILES` | `Files.Read.All` / `Files.ReadWrite.All` | 전 사용자 OneDrive 포함. 대상 드라이브 한정 검토 |
| `SEC-TEAMSRSC` | `Chat.*.All` / `ChannelMessage.*.All` | RSC(`*.Group`)로 특정 팀 한정 권고 |

#### C. 고위험 권한 — `Critical`

| 규칙 ID | 대상 권한 | 사유 |
|---|---|---|
| `RISK-PRIVESC` | `Directory.ReadWrite.All`, `RoleManagement.ReadWrite.*`, `Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`, `Policy.ReadWrite.*`, `PrivilegedAccess.*`, `GroupMember.ReadWrite.All`, `User.ReadWrite.All` | **권한 상승 경로**. 앱이 스스로 권한을 확대하거나 관리자 역할을 부여할 수 있음 → 정보보호그룹 필수 검토 |
| `RISK-IMPERSONATE` | `Mail.Send`(Application), `Mail.ReadWrite`(Application), `Chat.ReadWrite.All` | **사칭 발신 가능**. 임직원 명의 메일/채팅 발송 → 피싱 악용 위험 |
| `RISK-DELETE` | 작업유형 `X` | 삭제 권한 포함 → **별도 협의** 대상 (현업 부여 제외 원칙) |
| `RISK-UNKNOWN` | 사전 미등재 권한 | 공식 문서 링크와 함께 정보유형 확인 요청. 미확인 시 민감(S) 가정 |

#### D. 운영·보안 위생 — `Medium`

| 규칙 ID | 조건 | 권고 |
|---|---|---|
| `OPS-EXPIRY-LONG` | 만료일이 신청일 +1년 초과 또는 `무기한` | 최대 1년 이내 설정 후 갱신 심사 |
| `OPS-AUTH` | 항상 | Client Secret 대신 **Managed Identity** 또는 인증서 기반 인증 권고. Secret 사용 시 만료·회전 주기 명시 |
| `OPS-STORE-DETAIL` | `storage`가 저장 있음 | 저장 위치·보관 기간·암호화·파기 기준 명시 |
| `OPS-DIAGRAM` | 데이터 흐름도 이미지 미첨부 | 데이터 흐름도 첨부 권고 (외부 반출 경로 확인 목적) |
| `OPS-ENV` | 개발계/가동계 구분 서술 없음 | 개발계 선 검증 후 가동계 이관 권고 |
| `OPS-MONITOR` | 항상 | 분기별 사용량 모니터링 대상. 90일 이상 미호출 시 회수 안내 |

### FR-7. 향후 진행사항

| 최종 등급 | 절차 | 담당 | 예상 소요 |
|---|---|---|---|
| **G1** (개인 권한 정보) | ① 요건 확정 → ② DX운영그룹 주관 수행 → ③ (필요시) 정보보호그룹 공유 → ④ IR 결재 → ⑤ 권한 할당 → ⑥ 분기 모니터링 | DX운영그룹 | 1~3영업일 |
| **G2** (전사 일반 정보) | ① 요건 확정 → ② DX운영그룹 주관 → ③ **정보보호그룹 공유 및 보안의견 수렴** → ④ IR 결재 → ⑤ 권한 할당 → ⑥ 분기 모니터링 | DX운영 + 정보보호 | 3~5영업일 |
| **G3** (전사 민감 정보) | ① **보완사항 이행(보안 요건 충족)** → ② 정보보호그룹 검토 → ③ 운영의견 종합, **필요시 Committee 상정** → ④ 운영성 검토 → ⑤ IR 결재 → ⑥ 권한 할당 → ⑦ **집중 모니터링** | 정보보호 + DX운영 + (Committee) | 2주 이상 |

- `Blocker` 보완사항이 1건이라도 있으면 **"접수 보류 — 보완 후 재제출"** 을 최상단에 표시하고 위 절차는 참고로 내린다.
- `별도 협의(X)` 플래그가 있으면 등급과 무관하게 **"삭제 권한 별도 협의 필요"** 를 절차에 추가한다.
- 모든 등급 공통 후속: 권한 만료일 도래 시 재심사, 분기 전수 조사, 이상 호출 탐지 대상 등록.

### FR-8. 결과 출력

| ID | 요구사항 |
|---|---|
| FR-8.1 | 종합 등급 배지 + API별 등급 카드 (1차 → 최종 변화 표시) |
| FR-8.2 | 판정 근거 테이블: 조합 코드(예: `SDR`), 각 축의 판정값과 근거 |
| FR-8.3 | 보완 요청사항: 심각도별 정렬, 규칙 ID·대상 API·권고 조치 표시 |
| FR-8.4 | 향후 진행사항: 단계별 체크리스트 |
| FR-8.5 | **IR 회신문 복사** — 그대로 붙여넣을 수 있는 텍스트 |
| FR-8.6 | **JSON 내보내기** — 입력 필드 + 판정 결과 + 발동 규칙 ID 전체 |
| FR-8.7 | 추출 필드 편집 패널 — 수정 시 즉시 재판정 |
| FR-8.8 | 인쇄(PDF) 레이아웃 지원 |

---

## 4. API 사전 (내장 데이터)

슬라이드 7의 23건을 정본으로 하고 동일 패밀리로 확장한다. 형식:

```js
{ name: "Mail.Read", family: "메일", infoType: "S", note: "메일 본문·헤더 읽기", src: "PPT p7" }
```

### 4-1. 슬라이드 7 정본 (23건)

| 권한 | 분류 | PPT 등급 초안 | → infoType |
|---|---|---|---|
| `Mail.ReadWrite` | 메일 | G3 | **S** |
| `Mail.Read` | 메일 | G3 | **S** |
| `User.ReadWrite.All` | 사용자 | G2/G3 | **GS** |
| `User.Read.All` | 사용자 | G1/G2 | **G** |
| `Directory.ReadWrite.All` | 디렉터리 | G2/G3 | **GS** |
| `Directory.Read.All` | 디렉터리 | G2/G3 | **GS** |
| `Group.ReadWrite.All` | 그룹 | G2 | **G** |
| `Group.Read.All` | 그룹 | G1/G2 | **G** |
| `AuditLog.Read.All` | 보안 | G2 | **G** |
| `ThreatSubmission.Read.All` | 보안 | G2 | **G** |
| `RoleManagement.Read.All` | 보안 | G2/G3 | **GS** |
| `Team.ReadBasic.All` | Teams | G1/G2 | **G** |
| `Channel.ReadBasic.All` | Teams | G1/G2 | **G** |
| `ChannelMember.Read.All` | Teams | G1/G2 | **G** |
| `Chat.ReadBasic.All` | Teams | G3 | **S** |
| `Sites.ReadWrite.All` | SharePoint | G2 | **G** |
| `DeviceManagementConfiguration.ReadWrite.All` | Intune | G3 | **S** |
| `RoleManagement.Read.Directory` | 보안 | G3 | **S** |
| `DeviceManagementManagedDevices.Read.All` | Intune | G2/G3 | **GS** |
| `DeviceManagementApps.Read.All` | Intune | G2/G3 | **GS** |
| `DeviceManagementServiceConfig.Read.All` | Intune | G2/G3 | **GS** |
| `DeviceManagementRBAC.Read.All` | Intune | G2/G3 | **GS** |
| `DeviceManagementConfiguration.Read.All` | Intune | G2/G3 | **GS** |

### 4-2. 패밀리 확장 (사전 미등재 시 fallback)

| 리소스 접두 | infoType | 근거 |
|---|---|---|
| `Mail.`, `MailboxSettings.`, `MailboxFolder.` | **S** | 메일 본문·설정 = 민감정보 High |
| `Chat.`, `ChatMessage.`, `ChannelMessage.`, `TeamsAppInstallation.` | **S** | Teams 채팅 = 민감정보 High |
| `Calendars.`, `Contacts.` | **S** | 상세 일정 = 민감정보 Low |
| `Files.`, `Notes.` | **S** | 개인 OneDrive/OneNote 원문 포함 |
| `Reports.`, `SecurityEvents.`, `SecurityIncident.`, `IdentityRiskEvent.` | **S** | 접속 기록·보안 이벤트 |
| `Application.`, `AppRoleAssignment.`, `Policy.`, `PrivilegedAccess.`, `RoleEligibility.` | **S** | 권한 상승 경로 |
| `Sites.`, `Group.`, `GroupMember.`, `Team.`, `Channel.`, `ChannelMember.`, `User.Read`, `People.`, `Presence.`, `Organization.`, `LicenseAssignment.`, `Directory.Read` | **G** | 조직도·목록·라이선스 = 일반 정보 |
| `DeviceManagement`, `Device.`, `RoleManagement.` | **GS** | 활용 필드에 따라 상이 |
| 그 외 | **UNKNOWN → S 가정** | 보수적 기본값 (D4) |

> ⚠️ `Sites.*`를 `G`로 두는 것은 슬라이드 7의 `Sites.ReadWrite.All = G2` 판정을 정본으로 따른 것이다. 단, **SPO 문서 원본을 적재·학습에 사용하는 경우** 슬라이드 3의 "민감정보 High — SPO문서 원본"에 해당하므로 `ADJ` 단계에서 담당자 상향 검토를 안내한다.

---

## 5. 데이터 모델

```jsonc
{
  "meta": { "tool": "GAPS", "version": "1.0", "judgedAt": "2026-08-11T09:00:00+09:00",
            "sourceFile": "요청_OCR포털.pptx", "draft": true },
  "request": {
    "taskName": "OCR 포털 알림 메일링",
    "taskOverview": "...",
    "requester": "홍길동/DX TF",
    "developer": "ㅇㅇㅇ/메가존",
    "scopeVolume": "OCR포털 전체 사용자 / Daily 배치",
    "consent": "전원 동의/인지",
    "storage": "N/A",
    "remark": "",
    "expiry": "2027-01-31",
    "apis": [{
      "index": 1, "raw": "Mail.Read : 특정 사서함을 통한 시스템 알람 메일링 목표",
      "permission": "Mail.Read", "purpose": "특정 사서함을 통한 시스템 알람 메일링 목표",
      "data": "없음 (별도 포털 데이터의 메일 발송 기능만 활용)",
      "appId": "OCR-Portal-PRD-DP", "declaredPermType": "Application", "account": "특정 공유사서함"
    }]
  },
  "judgement": {
    "apis": [{
      "index": 1, "permission": "Mail.Read",
      "axes": {
        "infoType":   { "value": "S", "source": "DICT/PPT-p7", "note": "메일 본문·헤더" },
        "accessScope":{ "value": "D", "declared": "Application", "overridden": true,
                        "rule": "RULE-SCOPE-D",
                        "evidence": ["특정 사서함 — '특정 사서함을 통한 시스템 알람 메일링 목표'",
                                     "특정 공유사서함 — 활용계정/App"] },
        "opType":     { "value": "R", "source": "NAME-SUFFIX" }
      },
      "combo": "SDR",
      "primaryGrade": "G1",
      "adjustments": [],
      "finalGrade": "G1",
      "flags": ["SCOPE_OVERRIDDEN"]
    }],
    "overallGrade": "G1",
    "blockerCount": 1,
    "findings": [
      { "id": "SEC-SCOPE", "severity": "High", "target": "Mail.Read",
        "title": "Application 권한 범위 제한 필수",
        "detail": "...", "actions": ["RBAC for Applications로 대상 사서함 한정", "..."] },
      { "id": "REQ-EXPIRY", "severity": "Blocker", "target": "-", "title": "권한 만료일 미기재", "detail": "..." }
    ],
    "nextSteps": { "grade": "G1", "owner": "DX운영그룹", "eta": "1~3영업일", "steps": ["..."] }
  }
}
```

---

## 6. 화면 정의

```
┌──────────────────────────────────────────────────────────────┐
│  GAPS — Graph API 권한 등급 판정        [DRAFT 배지]  [도움말] │
├──────────────────────────────────────────────────────────────┤
│  ①  [ 신청 템플릿 PPTX를 여기에 끌어다 놓으세요 ]              │
│      · 파일은 브라우저에서만 처리되며 외부로 전송되지 않습니다  │
│      [파일 선택]   [샘플로 체험하기]                          │
├──────────────────────────────────────────────────────────────┤
│  ②  추출된 신청 정보                          [편집]          │
│      과제명 / 요청자 / API별 항목 …  (미기재 = 빨강 표시)      │
├──────────────────────────────────────────────────────────────┤
│  ③  판정 결과                                                 │
│      ┌───────────┐                                            │
│      │  종합 G1  │  API 1건 · 별도협의 없음                    │
│      └───────────┘                                            │
│      ▸ Mail.Read     S · D(재판정) · R  →  SDR  →  G1         │
│        1차 G3 ─(RULE-SCOPE-D)→ G1                             │
│        근거: "특정 사서함을 통한 시스템 알람 메일링 목표"       │
├──────────────────────────────────────────────────────────────┤
│  ④  보완 요청사항  (Blocker 1 · Critical 0 · High 1 · Med 3)  │
│      [Blocker] REQ-EXPIRY  권한 만료일 미기재 → 접수 보류      │
│      [High]    SEC-SCOPE   범위 제한 조치 필수 → 조치 목록…    │
├──────────────────────────────────────────────────────────────┤
│  ⑤  향후 진행사항                                             │
│      □ 보완사항 회신 → □ 요건 확정 → □ IR 결재 → □ 권한 할당   │
├──────────────────────────────────────────────────────────────┤
│      [IR 회신문 복사]  [JSON 내보내기]  [인쇄]                 │
└──────────────────────────────────────────────────────────────┘
```

**등급 색상**: G1 `#2C5F2D`(녹) · G2 `#B85042`(주황·경고) · G3 `#990011`(적) · Blocker `#212121`

---

## 7. 비기능 요구사항

| ID | 항목 | 기준 |
|---|---|---|
| NFR-1 | 배포 | **단일 HTML 파일**. 더블클릭 또는 사내 파일서버 링크로 실행 |
| NFR-2 | 네트워크 | 외부 요청 **0건**. 업로드 파일은 메모리에서만 처리, 저장·전송 없음 |
| NFR-3 | 브라우저 | Edge/Chrome(Chromium) 최신 2개 버전. `DecompressionStream` 미지원 시 내장 inflate 폴백 |
| NFR-4 | 성능 | 5MB PPTX 기준 파싱+판정 **2초 이내** |
| NFR-5 | 접근성 | 등급을 색상만으로 구분하지 않음(코드·아이콘 병기). 키보드 조작 가능 |
| NFR-6 | 유지보수 | API 사전·키워드·규칙을 파일 상단 **단일 설정 블록**으로 분리하여 비개발자도 수정 가능 |
| NFR-7 | 감사성 | 모든 판정에 규칙 ID 부여. JSON 출력에 발동 규칙 전량 포함 |
| NFR-8 | 다크모드 | OS 설정 자동 대응 |

---

## 8. 테스트 케이스 (수용 기준) — 검증 완료 2026-08-11

검증 방법: `samples/` 의 실제 PPTX를 브라우저에서 업로드 → 파싱 → 판정까지 전 구간 실행.

| TC | 샘플 파일 | 입력 요지 | 기대 결과 | 실측 | 결과 |
|---|---|---|---|---|---|
| **TC-1** | `샘플_TC1_특정사서함_알람메일.pptx` | `Mail.Read` / Application / "특정 사서함을 통한 시스템 알람 메일링" | `RULE-SCOPE-D` 발동 → `SDR`=**G1** + `SEC-SCOPE` High | `SDR` G1→G1, 근거 「특정 사서함」「공유사서함」, Blocker 0 / High 1 | ✅ |
| **TC-2** | `샘플_TC2_전사메일수집.pptx` | `Mail.Read`+`Sites.ReadWrite.All` / Application / 전사 수집·적재 | 재판정 미발동, 종합 **G3**, `RISK-MASSREAD` Critical | `SAR` G3, `GAW` G2→G3(`ADJ-SCOPE-UP`), 종합 **G3** | ✅ |
| **TC-3** | (TC-2에 포함) | `Sites.ReadWrite.All` / Application | `GAW`=**G2** + `SEC-SITES` | `GAW` base G2 (규모 상향으로 최종 G3) | ✅ |
| **TC-5** | `샘플_TC5_본문제외_미저장.pptx` | `Mail.ReadWrite` / Application / 본문 필드 제외 + 미저장 + 동의 확보 | `SAW`=G3 → **G2** (1등급 하향 캡) | `SAW` G3→**G2**, `ADJ-MIN`+`ADJ-CAP` | ✅ |
| **TC-6** | `샘플_TC6_임원Daily발송.pptx` | `User.Read.All` / Application / "직책자 지수를 임원에게 Daily 발송" | `GAR`=G2 → **G3** | `GAR` G2→**G3**, `ADJ-SCOPE-UP` | ✅ |
| **TC-7** | (TC-2) | 권한 만료일 미기재 | `REQ-EXPIRY` **Blocker** | Blocker 2건(`REQ-EXPIRY`, `REQ-CONSENT-NO`) | ✅ |
| **TC-8** | `샘플_TC8_디렉터리쓰기_퇴직자정리.pptx` | `Directory.ReadWrite.All` / Application | **G3** + `RISK-PRIVESC` Critical | `SAW` G3, Critical 2 (`RISK-PRIVESC`+`RISK-DELETE`) | ✅ |
| **TC-9** | (TC-2) | API 2건 등급 상이 | 종합 = 최고 등급 | `Mail.Read` G3 / `Sites.*` G3 → 종합 **G3** | ✅ |
| **TC-10** | (TC-8) | 활용목적에 "라이선스 제거" | 작업유형 **X** + `RISK-DELETE` + 별도 협의 플래그 | op=`X`, `RISK-DELETE` Critical, 플래그 표시 | ✅ |
| **TC-12** | (합성 바이트) | DRM(IRM) 암호화 PPTX | 파싱 전 차단 + 해제 안내 | OLE 시그니처 감지 → 안내 메시지 출력 | ✅ |
| **TC-13** | `샘플_TC6/TC8/TC14` (약식) | 약식 템플릿(슬라이드 11 구조) | Full과 동일하게 전 필드 추출 | `활용계정/App` 포함 전 필드 추출 | ✅ |
| **TC-14** | `샘플_TC14_범위상충.pptx` | Application / "특정 팀" + "전사 모든 팀" 동시 | 재판정 미발동 + `SEC-CONFLICT` | conflict=true, `SAW` **G3** 유지 | ✅ |
| **TC-15** | 빈 템플릿 | 배포용 빈 양식 업로드 | 전 필수항목 Blocker, 안내문(`*`)은 값으로 오인 금지 | Blocker 3건, 안내문 미수집 | ✅ |
| **TC-16** | 편집 재판정 | UI에서 "특정 사서함"→"전사 사서함"으로 수정 | 즉시 `SDR`/G1 → `SAR`/G3 전환 | 320ms 디바운스 후 G3 전환 | ✅ |
| **TC-17** | inflate 폴백 | `DecompressionStream` 제거 후 파싱 | 네이티브 결과와 동일 | 파싱 결과 **바이트 단위 일치** | ✅ |

| **TC-4** | (엔진 직접 호출) | `Group.Read.All` / Delegated | `GDR`=**G1** | `GDR` G1→G1, Blocker/Critical/High 0 | ✅ |
| **TC-11** | (엔진 직접 호출) | 사전 미등재 `FooBar.Read.All` / Application | `UNKNOWN`→S 가정 → 1차 **G3** + `RISK-UNKNOWN` | `SAR` base **G3** + `RISK-UNKNOWN`·`SEC-ALL` High<br>(입력의 저장=N/A로 `ADJ-STORE` 발동 → 최종 G2) | ✅ |
| **TC-18** | (엔진 직접 호출) | `Mail.ReadWrite` / 목적은 "제목 조회하여 대시보드 표시" | `SEC-OVERPERM` High | `SEC-OVERPERM` + `NOTE-RW-DELETE` 발동 | ✅ |
| **TC-19** | (엔진 직접 호출) | `Mail.ReadWrite` / 목적은 "알림 메일 발송" | `SEC-MAILSEND` High (Mail.Send 권고) | `SEC-MAILSEND` 발동 | ✅ |

---

## 9. 향후 확장 (v2+)

| 항목 | 내용 |
|---|---|
| E1 | 정보보호그룹 확정 기준으로 API 사전 교체 및 전 권한 확장 |
| E2 | IR 시스템 연동 — 판정 JSON 자동 첨부 |
| E3 | 요청 대장 자동 적재 (SharePoint 리스트) |
| E4 | 부여 권한 vs 실제 호출 로그 대조 → 미사용 권한 회수 대상 자동 산출 |
| E5 | 등급 분포·처리 리드타임 대시보드 |
| E6 | Excel(.xlsx) 신청서 포맷 병행 지원 |
