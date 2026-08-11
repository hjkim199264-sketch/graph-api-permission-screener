# M365 권한 요청 대응 체계 — Day2 산출물

DX혁신실 M365 운영 | 2026-08-11

두 갈래의 권한 요청을 같은 판정 구조로 처리하기 위한 기준·문서·도구 모음이다.

| | 대상 | 등급 | 입력 |
|---|---|---|---|
| **Tab 1** | Graph API 권한 (사전 심사) | G1 / G2 / G3 | 신청 템플릿 PPTX 업로드 |
| **Tab 2** | M365 관리자 역할 (사전 심사) | A1 / A2 / A3 | 역할 검색 또는 목록 선택 |
| **Tab 3** | Entra App 회수 (사후 정리) | D1 / D2 / D3 / D4 | 인벤토리 CSV 업로드 |
| **Tab 4** | 라이선스 회수 (사후 정리) | L1 / L2 / L3 / L4 | 계정 · 사용량 CSV 업로드 |

두 체계 모두 **3기준 × 2값 = 8조합 → 3등급** 이라는 동일한 구조를 쓴다.
Graph API는 "Delegated면 무조건 G1", 관리자 역할은 "민감 자산 접근이 있으면 무조건 A3" 처럼
**한 축이 등급을 결정**하는 방식도 같다.

---

## 바로 쓰기

### 웹으로 (팀 공유용)

**https://graph-api-permission-screener.vercel.app**

### 파일로 (오프라인·사내망)

1. **`GraphAPI_등급판정시스템.html`** 을 더블클릭 (Edge/Chrome)
2. `samples/` 안의 PPTX를 창에 끌어다 놓기 — 또는 **[샘플로 체험하기]** 클릭
3. 결과 확인 후 **[IR 회신문 복사]** 로 그대로 회신

> 어느 쪽이든 **업로드한 파일은 브라우저 안에서만** 처리된다. 서버로 전송되지 않으며, 웹 버전도 정적 호스팅일 뿐 파일을 수신·보관하지 않는다.
> 파일 버전은 설치·인터넷 연결도 필요 없다.

- 저장소 · https://github.com/hjkim199264-sketch/graph-api-permission-screener
- 배포 URL은 검색엔진 색인에서 제외(`X-Robots-Tag: noindex`)했으나, **링크를 아는 사람은 접근할 수 있다.**

---

## 파일 구성

| 경로 | 설명 |
|---|---|
| `GraphAPI_등급판정시스템.html` | **판정 시스템 본체 (Tab1 + Tab2).** 단일 파일·오프라인 동작 |
| `index.html` | 웹 배포용 진입점 (위 파일로 리다이렉트) |
| **문서** | |
| `docs/BRD_M365_권한관리체계.md` | 배경·문제정의·목표·범위·등급 체계 3종·성공지표·리스크·로드맵 |
| `docs/PRD_GAPS_통합판정시스템.md` | 3탭 기능요구사항·판정 규칙·데이터 모델·화면·테스트 결과 전량 |
| `docs/HISTORY.md` | **작업 이력** — 타임라인·설계 결정과 근거·발견한 이슈·검증 이력 |
| **Graph API (Tab 1)** | |
| `templates/Graph API 권한 신청서_템플릿_v1.0.pptx` | **요청자 배포용 빈 양식** (Full 1p + 약식 1p) |
| `samples/*.pptx` | 작성 예시 6종 (테스트 케이스 겸용) |
| `tools/build_templates.py` | 위 템플릿·샘플 PPTX 생성 스크립트 |
| **관리자 역할 (Tab 2)** | |
| `docs/M365 관리자 권한 분류 및 등급 기준_v1.0.pptx` | **기준 장표 14p** — 배경, 3기준×2값, 8조합→A1/A2/A3, 승인 프로세스, 역할별 등급표 |
| `docs/M365 관리자 역할 명세서_v1.0.xlsx` | **역할 명세 136건** — 권한/위험도/핵심 리스크 + 등급·3축·보안검토·PIM·공식문서 (3시트) |
| `tools/admin_roles_data.py` | **역할 데이터 단일 원본.** 기준을 바꾸려면 이 파일만 수정 |
| `tools/build_admin_role_docs.py` | 위 데이터로 PPT·Excel·JSON 생성 + HTML에 데이터 주입 |
| `tools/admin_roles.json` | 생성된 역할 데이터 (HTML 임베드용) |
| **Entra App 회수 (Tab 3)** | |
| `tools/collect_entra_inventory.ps1` | 관리센터 CSV에 없는 항목(Graph 권한·소유자·동의 내역·로그인 활동)을 Graph에서 수집 |
| `docs/모니터링_데이터수집_가이드.md` | CSV 실측 분석 + 수집 방법 + Graph 활동 로그 설정 + KQL 쿼리 |
| **공통** | |
| `vercel.json` | 배포 설정 (보안 헤더, `noindex`) |

> ⚠ 실제 테넌트 CSV(`AppRegistrationList*.csv`, `entra_export/`, `0N_*.csv`)는 `.gitignore`에 등록되어 있다.
> 저장소가 public이므로 실수로 커밋되지 않도록 반드시 유지할 것.

## 업데이트 배포

```bash
python tools/build_admin_role_docs.py && git add -A && git commit -m "..." && git push && vercel deploy --prod --yes
```

---

## Tab 3 — Entra App 회수 대상 산출 (D1~D4)

Tab 1·2가 **사전 심사**라면 Tab 3은 **사후 정리**다. 인벤토리 CSV를 올리면 처분 등급과
소유자별 확인 메일 초안까지 산출한다.

### 두 가지 모드

| 모드 | 입력 | 판정 가능 범위 |
|---|---|---|
| **기본** | M365 관리센터 &gt; 앱 등록 &gt; **관리 목록 내보내기** CSV 1개 | 자격증명·리디렉션 URI·태그·멀티테넌트 기반 D1/D2/D4 |
| **정밀** | + `collect_entra_inventory.ps1` 산출 CSV (01~05) | Graph 권한·소유자·로그인 활동 반영 → **D3 판정 + 소유자별 메일** |

화면 ②에 **데이터 커버리지**가 표시되어, 무엇이 없어서 무엇을 못 하는지 항상 드러난다.

### 처분 등급

| | 판정 | 조건 |
|---|---|---|
| **D1** | 즉시 삭제 후보 | 자격증명·리디렉션 URI·Graph 권한이 모두 없고 방치 기준(기본 180일) 경과 |
| **D2** | 소유자 확인 필요 | 자격증명이 없거나 만료됐으나 구성 요소가 남아 있음 |
| **D3** | 권한 회수 대상 | Graph 권한 보유 + 사용 이력 없음 또는 인증 불가 |
| **D4** | 유지 | 유효 자격증명 + 최근 사용 |

**보안 플래그**는 처분과 별개로 병행 표시한다 —
`MT` 멀티테넌트 · `PA` 개인계정 허용 · `IMP` 암묵적 허용 흐름 · `EXP` 90일내 만료 · `DEAD` 전부 만료 · `NOOWNER` 소유자 없음

### 소유자별 확인 메일

D1·D2·D3 대상을 소유자별로 묶어 메일 초안을 만든다. 회신 옵션(사용 중 / 미사용 / 판단 어려움),
회신 기한(10영업일), 미회신 시 처리 절차가 포함된다.

**자동 생성 앱**(Copilot Studio 등)과 **소유자 미지정 앱**은 개별 메일 대상에서 제외하고 별도 안내한다.
전자는 "만든 적 없다"는 회신이 돌아올 가능성이 높아 공지 방식이 적합하고, 후자는 보낼 대상이 없다.

> 실측 참고 — 사내 앱 등록 499건 분석 시 **자동 생성 앱이 323건(65%)** 이었다.
> 자세한 내용은 [모니터링_데이터수집_가이드.md](docs/모니터링_데이터수집_가이드.md).

---

## Tab 4 — 라이선스 회수 대상 산출 (L1~L4)

미사용 라이선스를 회수해야 할 사람을 찾는다. 판정 결과에 절감 추정과 본인 앞 통보 메일 초안이 따라온다.

### 입력 CSV 3종

| | 파일 | 경로 | 없으면 |
|---|---|---|---|
| ① | **계정 현황** (필수) | 관리센터 &gt; 사용자 &gt; 활성 사용자 &gt; **사용자 내보내기** | 아무것도 못 함 |
| ② | **활성 사용자 - 세부 정보** (권장) | 관리센터 &gt; 보고서 &gt; 사용량 | **"미사용" 판정 불가.** 구조적 낭비만 잡힘 |
| ③ | **Microsoft 365 Copilot** (선택) | 관리센터 &gt; 보고서 &gt; 사용량 | Copilot 미사용자 식별 불가 |

> ⚠ **먼저 확인** — 관리센터 &gt; 설정 &gt; 조직 설정 &gt; 보고서에서 **「익명 식별자 표시」가 켜져 있으면**
> 사용자 이름이 해시로 나와 계정 CSV와 매칭되지 않는다. 내보내기 전에 꺼야 한다.
> 도구가 자동 감지해 경고하지만, 애초에 꺼두는 편이 낫다.

### 판정 등급

| | 판정 | 조건 |
|---|---|---|
| **L1** | 즉시 회수 | 차단·삭제 계정 보유 · 게스트 보유 · 할당 후 활동 이력 전무 |
| **L2** | 회수 검토 | 기준 기간(기본 90일) 이상 전 서비스 활동 없음 |
| **L3** | 다운그레이드 검토 | Copilot 미사용 · 상하위 SKU 중복 · Exchange만 사용 |
| **L4** | 유지 | 정상 사용 중 |

플래그 — `BLOCK` `DELETED` `GUEST` `NEVER` `DUP` `NOLOC` `NOUSAGE` `NOCOPILOT` `MAILONLY`

### 절감 추정

SKU별 단가를 직접 입력하면 연간 절감액이 계산된다. 계약 단가는 EA/CSP 등 형태마다 다르므로
기본값을 넣지 않는다. L3는 상한값이며 실제로는 하위 SKU 비용이 차감된다.

### 통보 메일

L1~L3 대상에게 보낼 초안을 만든다. 회신 3옵션(계속 필요 / 불필요 / 조정 동의), 회신 기한 10영업일,
미회신 시 처리, **라이선스 회수 후 30일 경과 시 사서함 데이터 삭제 가능** 경고가 포함된다.

차단·게스트 계정은 통보 대상에서 제외하되, **퇴직 처리 중이거나 복직 예정 계정이 섞일 수 있으니
인사 담당과 대조 후 진행**하라고 안내한다.

> 실측 참고 — 계정 5,427건(게스트 1,457 포함) 중 유상 라이선스 보유 3,548명.
> 계정 CSV만으로는 회수·조정 대상이 30명(1%)뿐이다. **사용량 리포트가 있어야 본론이 시작된다.**

---

## Tab 2 — M365 관리자 역할 등급 (A1/A2/A3)

Microsoft Entra 기본 제공 역할 **136종**을 판정한 결과가 내장되어 있다.
역할명(한글/영문)으로 검색하거나 목록에서 고르면 등급·위험도·핵심 리스크·보안 검토 필요 여부·승인 절차가 나온다.

### 3가지 판정 기준

| 기준 | 값 |
|---|---|
| ① 영향 범위 (Scope) | **S** 서비스·기능 한정 / **T** 테넌트 전역 |
| ② 작업 유형 (Operation) | **R** 읽기 전용 / **W** 구성 변경 |
| ③ 민감 자산 접근 (Sensitive) | **N** 없음 / **P** 있음 |

**민감 자산 접근(P)** 은 다음 두 경로 중 하나라도 해당할 때다.

- **(a) 권한 상승** — 역할 할당, 자격 증명·MFA 변경, 앱 권한 동의, 조건부 액세스·인증 정책 변경, 도메인·페더레이션 변경
- **(b) 콘텐츠 접근** — 메일·파일·채팅·문서 원문 열람 또는 열람 위임 설정 가능

### 8조합 → 3등급

```
민감 접근 P 있음                  →  A3   (범위·작업과 무관)
N + 읽기(R) + 서비스(S)           →  A1
N + 읽기(R) + 전역(T)             →  A2
N + 변경(W) + 서비스(S)           →  A2
N + 변경(W) + 전역(T)             →  A3
```

판정 결과: **A3 65건 (47%) · A2 58건 (42%) · A1 13건 (9%)**

A3 비중이 높은 것은 관리자 역할 다수가 구조적으로 권한 상승 또는 콘텐츠 접근 경로를 갖기 때문이며,
관리자 역할 요청을 Graph API보다 엄격히 통제해야 하는 근거가 된다.

### 등급별 차등 적용

| | A1 🟢 낮음 | A2 🟡 중간 | A3 🔴 높음 |
|---|---|---|---|
| 보안 검토 | 생략 (사후 이력) | 정보보호그룹 공유·의견 수렴 | **검토 필수** (승인 없이는 부여 불가) |
| 할당 방식 | 상시 할당 가능 | PIM 적격 할당 권장 | **PIM 필수 · 상시 할당 금지** |
| 활성화 조건 | — | MFA | MFA + 승인자 승인 + 사유 기록 (최대 8시간) |
| 할당 기간 | 1년 | 6개월 | 3개월 |
| 계정 분리 | 일반 계정 가능 | 일반 계정 가능 | **관리자 전용 계정 필수** |
| 예상 소요 | 1영업일 | 3~5영업일 | 2주 이상 |

> ★ 표시는 Microsoft 공식 「권한 있는 역할(Privileged role)」 — 136건 중 36건.
> 이는 사내 등급(A1/A2/A3)과 별개의 Microsoft 자체 표시이며, 참고 지표로 함께 노출한다.

---

## Tab 1 — Graph API 권한 등급 (G1/G2/G3)

### 1차 분류 — 3기준 × 2값 = 8조합

| 기준 | 값 |
|---|---|
| ① 정보 유형 | **G** 일반 / **S** 민감 |
| ② 접근 범위 | **D** 개인·Delegated / **A** 전사·Application |
| ③ 작업 유형 | **R** Read / **W** Create·Update / *(X Delete → 별도 협의)* |

```
GDR · GDW · SDR · SDW  →  G1  (개인 권한 정보)
GAR · GAW              →  G2  (전사 일반 정보)
SAR · SAW              →  G3  (전사 민감 정보)
```

즉 **접근 범위가 Delegated면 정보 민감도와 무관하게 G1**, Application일 때만 정보 유형이 G2/G3를 가른다.

### ★ 접근범위 재판정 (RULE-SCOPE-D)

**Application 권한으로 신청되었더라도, 「특정 사서함」·「특정 팀」·「공유 사서함」·「로그인한 사용자」 등
대상 한정 서술이 있으면 접근범위를 Delegated로 재판정**한다. (「전사」·「전체 사용자」 등 전사 표현이 함께
있으면 재판정하지 않고 *범위 상충* 으로 표시)

재판정으로 낮아진 등급은 **기술적 범위 제한 조치의 이행을 전제 조건**으로 하며, 미이행 시 원 등급으로 환원한다.

| 워크로드 | 조치 |
|---|---|
| Exchange Online | RBAC for Applications(권장) 또는 `New-ApplicationAccessPolicy` |
| SharePoint/OneDrive | `Sites.Selected` + 대상 사이트 지정 |
| Teams | RSC(Resource-Specific Consent) / `*.Group` 계열 권한 |
| 공통 | Delegated 권한으로 전환 (최우선 대안) |

> 근거: Application 권한은 별도 제한이 없으면 **테넌트 전체**에 적용된다.
> "특정 사서함만 쓰겠다"는 선언은 코드 레벨의 약속일 뿐, 부여된 권한 자체를 좁히지 않는다.

### 2차 분류 — 4개 판단 지표

| 지표 | 조정 |
|---|---|
| 데이터 최소화(필터링/마스킹) | G3 → G2 |
| 정보주체 인지/동의 | G2 → G1 |
| 저장·2차 가공 (비저장) | G3 → G2 |
| 영향 범위/활용 규모 | G2 → G3 (상향) |

**하향은 최대 1등급**, 상향·하향 동시 발동 시 상쇄(1차 등급 유지). 모든 조정은 규칙 ID와 근거를 함께 출력한다.

---

## 기준을 바꾸려면

### Tab 1 (Graph API)

HTML 파일 상단의 `const CFG = { … }` **설정 블록만** 수정하면 된다.

| 항목 | 용도 |
|---|---|
| `CFG.DICT` | API 사전 — 권한별 정보유형(G/S/GS). 슬라이드 7의 23건이 정본 |
| `CFG.FAMILY` | 사전 미등재 권한의 정보유형 추정 규칙 |
| `CFG.KW_LIMIT` / `KW_TENANT` | 재판정 키워드 — 대상 한정 표현 / 전사 표현 |
| `CFG.RE_*` | 2차 분류 판정 정규식 (최소화·동의·저장·규모) |
| `CFG.PRIVESC` / `IMPERSONATE` / `MASSREAD` | 고위험 권한 목록 (Critical 판정) |
| `CFG.GRADE_INFO` / `CFG.STEPS` | 등급별 절차 — 담당 조직·소요·단계 |

정보보호그룹의 민감정보 기준이 확정되면 `CFG.DICT` 의 `info` 값을 교체한다.

### Tab 2 (관리자 역할)

HTML을 직접 고치지 말고 **`tools/admin_roles_data.py` 만 수정한 뒤 빌드**한다.
이 파일 하나가 PPT·Excel·HTML 3종의 공통 원본이다.

```bash
python tools/build_admin_role_docs.py
```

| 항목 | 용도 |
|---|---|
| `ROLES` | 역할별 3축 판정 + 범주 + 핵심 리스크 (136건) |
| `grade_of()` | 8조합 → A1/A2/A3 매핑 함수 |
| `GRADE_DEF` | 등급별 정의·승인 절차·보안 검토·PIM·소요 |
| `SENS_REASON` | 민감 자산 접근(P) 판정 사유 코드 |

HTML의 `/* ADMIN_ROLES_DATA_START */ … /* ADMIN_ROLES_DATA_END */` 블록은
빌드 스크립트가 자동으로 덮어쓰므로 직접 수정하지 않는다.

---

## 알아둘 점

- **판정 결과는 초안(Draft)** 이다. 최종 승인은 정보보호그룹 검토와 IR 결재선을 따른다.
- Tab 1에서 파싱이 틀리면 **② 추출된 신청 정보** 에서 직접 고칠 수 있고, 수정 즉시 재판정된다.
- **DRM/IRM(민감도 레이블) 암호화 PPTX는 읽을 수 없다.** PowerPoint에서 열어 레이블을 해제한 뒤 올려야 한다.
- 레거시 `.ppt` 는 미지원 — `.pptx` 로 저장 후 사용.
- Graph API 사전에 없는 권한은 **보수적으로 민감(S)** 으로 가정하고 확인을 요청한다.
- Tab 2의 역할 목록은 Microsoft Entra 기본 제공 역할 기준이며, MS가 역할을 추가·변경하면
  `admin_roles_data.py` 를 갱신해야 한다. Intune 자체 RBAC 역할과 Purview 내부 역할은 포함하지 않는다.

## 참고

**Graph API**
- [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Limiting application permissions to specific mailboxes](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
- [Sites.Selected 개요](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)

**관리자 역할**
- [Microsoft 365 관리 센터 관리자 역할 정보](https://learn.microsoft.com/ko-kr/microsoft-365/admin/add-users/about-admin-roles)
- [Microsoft Entra 기본 제공 역할](https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/permissions-reference)
- [권한 있는 역할 및 사용 권한](https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/privileged-roles-permissions)
- [태스크별 최소 권한 있는 역할](https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/delegate-by-task)
