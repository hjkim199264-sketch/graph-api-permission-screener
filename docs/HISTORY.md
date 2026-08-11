# 작업 이력 (HISTORY)

M365 권한 거버넌스 체계 구축 — DX혁신실 M365 운영
작업일: 2026-08-11 · 저장소: https://github.com/hjkim199264-sketch/graph-api-permission-screener
배포: https://graph-api-permission-screener.vercel.app

---

## 0. 요약

「Graph API 할당 프로세스 및 기준_260806」 PPT의 등급 분류 기준을 출발점으로,
**신청 → 판정 → 승인 → 사후 회수** 전 구간을 지원하는 문서·도구 일체를 구축했다.

| | 구간 | 대상 | 등급 | 입력 |
|---|---|---|---|---|
| Tab 1 | 사전 심사 | Graph API 권한 | G1 / G2 / G3 | 신청 템플릿 PPTX |
| Tab 2 | 사전 심사 | M365 관리자 역할 | A1 / A2 / A3 | 역할 검색·선택 |
| Tab 3 | 사후 정리 | Entra App 회수 | D1 / D2 / D3 / D4 | 인벤토리 CSV |

세 체계 모두 **3기준 × 2값 = 8조합 → 3등급** 이라는 같은 골격을 쓰고,
**한 축이 등급을 결정**하는 구조(아래 2-2)도 공유한다.

---

## 1. 타임라인

### 1-1. 원본 PPT 분석

원본 `Graph API 할당 프로세스 및 기준_260806_학습용.pptx`(11장)를 읽는 데 실패했다.

- 파일 선두가 `D0 CF 11 E0`(OLE Compound File) — ZIP이 아니었다
- 내부 스트림에 `DRMEncryptedDataSpace`, `Microsoft.Metadata.DRMTransform`, `EncryptedPackage` 확인
- → **Microsoft Purview 민감도 레이블 암호화** 적용 상태

PowerShell COM으로 PowerPoint를 통해 여는 우회를 시도했으나 사용자가 직접 암호를 해제해 주어 해결.
이 경험은 그대로 제품 요구사항이 되었다 — **판정 도구도 DRM 파일을 열 수 없으므로 사전 감지 후 안내**(FR-1.3).

추출 결과 확보한 핵심 기준:

- 3기준 × 2값 = 8조합 → G1/G2/G3 (슬라이드 3·4)
- 2차 분류 4개 판단 지표 (슬라이드 4)
- 요청 템플릿 Full / 약식 2종 (슬라이드 5·11)
- 주요 API 등급 분류 초안 23건 (슬라이드 7)
- 처리 이력 12건 (슬라이드 9) — 접근범위 재판정 룰의 근거

### 1-2. BRD / PRD 작성

`docs/BRD_M365_권한관리체계.md`, `docs/PRD_GAPS_통합판정시스템.md` 초안 작성.
Microsoft Graph 공식 문서로 Delegated / Application 권한 모델을 검증했다.

### 1-3. 신청 템플릿 · 테스트 샘플 생성

`tools/build_templates.py` — 원본 슬라이드 5·11 표 구조를 그대로 재현.

- 배포용 빈 양식 1종 (Full 1p + 약식 1p)
- 테스트 겸 작성 예시 6종 (TC-1·2·5·6·8·14)

생성 후 표 구조(`gridSpan` 7개, `hMerge` 14개)가 원본과 동일한지 확인.

### 1-4. Tab 1 — Graph API 판정 시스템

`GraphAPI_등급판정시스템.html` 단일 파일 구현. 외부 CDN 0건, 네트워크 요청 0건.

핵심 구현:

- **raw DEFLATE 압축 해제** — 네이티브 `DecompressionStream('deflate-raw')` 사용,
  미지원 브라우저용 puff 방식 inflate를 직접 구현해 폴백
- ZIP 중앙 디렉터리 파싱 → `ppt/slides/slideN.xml` 추출
- 표(`a:tbl`) **라벨 기반** 필드 매핑 — 셀 위치가 아닌 라벨 문구로 찾아 Full/약식 양쪽 대응
- 판정 엔진 (1차 8조합 → 2차 4지표 → 보완 요청사항 → 향후 진행사항)

**E2E 검증에서 실제 버그 1건 발견·수정** — 아래 3-1.

### 1-5. GitHub / Vercel 배포

공개 여부를 사용자에게 확인 후 진행. Public 저장소 + Vercel Hobby 배포.
검색엔진 색인은 `X-Robots-Tag: noindex, nofollow`로 차단.

`vercel git connect`는 Vercel GitHub App 미설치로 실패 → CLI 수동 배포로 운영.

### 1-6. UI 개선

- 업로드 문구를 "Graph API 신청서"로 명확화
- 메인에 **기준 요약 섹션** 추가 (3기준 / 8조합 매핑 / 재판정 룰 / 2차 분류 / 작성 유의사항)
- 업로드 영역을 최상단으로 이동, 할당 프로세스 5단계 섹션 제거 (사용자 요청)
- 결과가 나오면 기준 요약 자동 접힘

### 1-7. Tab 2 — M365 관리자 역할 등급 체계

Microsoft Entra 기본 제공 역할 **109종 + M365 관리센터 노출 역할 = 136종**을 판정.
MS 공식 「권한 있는 역할」 36건 표시를 별도 지표로 병기.

Graph API 체계와 **의도적으로 같은 구조**로 설계 (3기준 × 2값 = 8조합 → 3등급).

산출물 3종을 단일 원본(`tools/admin_roles_data.py`)에서 생성:

- `docs/M365 관리자 권한 분류 및 등급 기준_v1.0.pptx` (14장)
- `docs/M365 관리자 역할 명세서_v1.0.xlsx` (3시트 136행)
- HTML Tab 2 임베드 데이터 (자동 주입)

PPT는 PowerPoint COM으로 전 장을 PNG 렌더링해 육안 검수 → 표 넘침 2건 수정 (아래 3-2).

### 1-8. 보안검토 요청 메일 (Tab 1)

판정 결과로 정보보호그룹 앞 메일 제목·본문 자동 작성.
등급별 분기 — G3 보안검토 요청(10영업일) / G2 보안의견 수렴(5영업일) / G1 부여 사전 공유.

**한글 본문의 mailto 길이 문제**를 발견해 설계 변경 (아래 3-3).

### 1-9. 앱 등록 CSV 실사 · 수집 스크립트

사용자가 M365 관리센터에서 내보낸 `AppRegistrationList (1).csv`(499건) 전수 분석.

**결론: 이 CSV로는 "Graph API 권한 없는 깡통 App"을 뽑을 수 없다.**
`requiredResourceAccess`(권한 목록), `owners`(소유자), `signInActivity`(사용 이력)가 모두 빠져 있다.

대신 다른 축에서 강한 신호를 찾았다 — 아래 4장.

`tools/collect_entra_inventory.ps1` — 빠진 항목을 Graph에서 직접 수집.
`docs/모니터링_데이터수집_가이드.md` — 실측 분석 + 수집 방법 + Graph 활동 로그 설정 + KQL.

### 1-10. Tab 3 — Entra App 회수 대상 산출기

인벤토리 CSV → 처분 등급(D1~D4) + 보안 플래그 + 소유자별 확인 메일 초안.

**두 가지 모드**로 설계 — 수집 스크립트 실행 전에도 지금 가진 CSV만으로 동작한다.
데이터 커버리지 패널로 "무엇이 없어서 무엇을 못 하는지"를 항상 노출한다.

---

## 2. 주요 설계 결정과 근거

### 2-1. 세 체계를 같은 골격으로 통일

Graph API 기준이 이미 「3기준 × 2값 = 8조합 → 3등급」이었으므로,
관리자 역할(A1~A3)도 같은 골격으로 설계했다.

- 판정 방식이 하나면 담당자가 두 번 배우지 않는다
- 요청자에게 설명하는 논리가 같다
- 화면·문서·코드 구조를 재사용할 수 있다

| | ① | ② | ③ |
|---|---|---|---|
| Graph API | 정보유형 G/S | 접근범위 D/A | 작업유형 R/W(/X) |
| 관리자 역할 | 영향범위 S/T | 작업유형 R/W | 민감자산접근 N/P |

### 2-2. "한 축이 등급을 결정한다"

두 체계 모두 하나의 축이 다른 축을 압도하도록 설계했다.

- **Graph API**: 접근범위가 **Delegated면 정보 민감도와 무관하게 G1**
  → 앱이 사용자 본인의 권한 범위를 넘을 수 없기 때문
- **관리자 역할**: **민감 자산 접근(P)이 있으면 범위·작업과 무관하게 A3**
  → 권한 상승과 콘텐츠 열람은 읽기/쓰기·범위와 무관하게 최고 위험이기 때문

전자는 원본 PPT의 매핑에서 도출했고, 후자는 그 논리를 관리자 역할에 적용한 것이다.

### 2-3. 접근범위 재판정 룰 (RULE-SCOPE-D)

사용자가 명시적으로 요청한 규칙 — *"Application 권한을 신청했지만 특정 사서함이나 특정 팀에
접근하겠다는 멘트가 있으면 Delegated로 분류"*.

처리 이력(슬라이드 9)에서 이미 반복되던 실무 패턴이었다.

| 이력 | 실제 처리 |
|---|---|
| 3번 | "현업 담당자 **1인의 메일 사서함 및 SharePoint에만** 접근 가능한 권한으로 제한하여 부여" |
| 8번 | "**로그인 된 특정 사용자의** 메일, SPO만 접근 가능한 제한적 권한으로 승인" |
| 11번 | "전체 열람 권한이 아닌, **로그인한 사용자가 열람 권한이 있는 파일만** 열람 가능하도록 제한" |

**여기에 안전장치를 하나 추가했다.** Application 권한은 별도 제한이 없으면 실제로는
테넌트 전체에 적용된다. "특정 사서함만 쓰겠다"는 선언은 코드 레벨의 약속일 뿐 부여된 권한을
좁히지 않는다. 그래서 재판정으로 낮아진 등급에 **기술적 범위 제한 조치 이행을 전제 조건으로 부착**하고,
미이행 시 원 등급 환원을 명시한다.

| 워크로드 | 조치 |
|---|---|
| Exchange Online | RBAC for Applications(권장) 또는 `New-ApplicationAccessPolicy` |
| SharePoint/OneDrive | `Sites.Selected` + 대상 사이트 지정 |
| Teams | RSC(Resource-Specific Consent) / `*.Group` 계열 |
| 공통 | Delegated 권한 전환 (최우선 대안) |

또한 「전사」·「전체 사용자」 같은 전사 표현이 함께 있으면 재판정하지 않고 **범위 상충**으로 표시한다.

### 2-4. 2차 분류 하향은 최대 1등급

`ADJ-MIN`(데이터 최소화)과 `ADJ-STORE`(비저장)가 동시 발동해도 G3 → G2까지만,
여기에 `ADJ-CONSENT`가 더해져도 G1로 내려가지 않는다.

2단계 하향은 "전사 민감"을 "개인 권한"과 동일 취급하는 것이라 위험하다. 담당자 재량으로만 허용한다.
상향(`ADJ-SCOPE-UP`)은 안전 방향이므로 제한하지 않는다.

### 2-5. 보수적 기본값

- 사전 미등재 Graph 권한 → **민감(S) 가정** + 확인 요청
- 권한 유형 미기재 → **Application 가정** (보수적)
- `G2/G3` 병기 권한 → **S 가정** + 활용 필드 확인 요청
- 미기재 항목은 **유리한 방향으로 해석하지 않는다** (미기재 ≠ 조건 충족)

### 2-6. 오프라인 단일 파일

신청서에 개인정보·기밀이 포함될 수 있어 **업로드 파일이 외부로 나가면 안 된다**.
사내 PC에서 설치 없이 돌아야 한다.

→ 단일 HTML, 외부 CDN 0건, 네트워크 요청 0건.
ZIP 해제는 네이티브 `DecompressionStream` + 자체 inflate 폴백(결과 일치 검증 완료).

### 2-7. 데이터 단일 원본

관리자 역할 데이터는 `tools/admin_roles_data.py` 하나가 PPT·Excel·HTML 3종의 원본이다.
빌드 스크립트가 HTML의 `/* ADMIN_ROLES_DATA_START */ … END */` 블록을 자동으로 덮어쓴다.

기준이 바뀌면 파일 하나만 고치고 빌드하면 세 산출물이 동시에 갱신된다.

### 2-8. 자동 판정은 초안일 뿐

- 모든 화면·출력물에 **DRAFT 배지**
- 모든 판정에 **발동 규칙 ID + 근거 원문** 병기
- 자동 승인 경로 없음. G3/A3는 정보보호그룹 승인 없이는 부여 불가
- 파싱 결과를 화면에서 **수기 보정** 가능 (수정 즉시 재판정)

---

## 3. 발견한 이슈와 해결

### 3-1. `저장/2차 가공: N/A`가 미기재로 처리된 버그 (실제 버그)

E2E 테스트 중 TC-1에서 잘못된 Blocker가 발생했다.

원인 — 공통 공백 판정 정규식(`RE_BLANK`)에 `n/a`가 포함되어 있어
`저장/2차 가공: N/A`가 "미기재"로 처리됐다. 그러나 이 필드에서 `N/A`는 **비저장**이라는
의미 있는 값이고, 2차 분류 하향의 근거가 된다.

해결 — `consent`/`storage` 전용 `RE_BLANK_STRICT`를 분리했다.
추가로 `consent: N/A`(동의 미확보)는 최종 등급에 따라 심각도를 달리하도록 개선
(G2 이상 `Blocker` / G1 `High`).

### 3-2. PPT 표 슬라이드 밖 넘침 (2건)

PowerPoint COM으로 14장을 PNG 렌더링해 육안 검수한 결과:

- 슬라이드 2 — 민감 사유 코드 표 8행이 슬라이드 하단을 넘어 각주와 겹침
  → 2열 레이아웃(4행 × 2쌍)으로 재구성
- 슬라이드 3 — 등급별 정의 표가 각주와 충돌
  → 「등급 / 구분」 열을 병합해 정의 열을 넓히고 위치 조정

역할 목록 슬라이드는 페이지당 17행 → 15행으로 낮춰 여유를 확보했다.

### 3-3. 한글 본문의 mailto 길이 한계

한글은 URL 인코딩 시 글자당 9자(`%XX%XX%XX`)로 늘어난다.
본문 2,355자 → 인코딩 11,999자. mailto 실질 한도(약 2,000자)를 크게 넘는다.

초기 구현은 "길면 제목만 채운다"였는데, 그러면 초안이 사실상 쓸모없다.

해결 — **전체 본문은 항상 클립보드에 복사**하고, 초안에는 요약(과제명·요청자·등급·API·검토 건수)과
"클립보드에서 붙여넣으세요" 안내를 채운다.

### 3-4. 검색 결과에 본문만 매칭된 역할이 섞임

Tab 2에서 "암호"로 검색하면 역할명과 무관하게 설명에 "암호화"가 들어간 역할까지 상위에 나왔다.

해결 — 매칭 위치별 랭킹 도입.
`이름 시작(0) > 한글명 포함(1) > 영문명 포함(2) > 범주(3) > 설명(4)` 순으로 정렬한다.

### 3-5. 관리센터 CSV 파싱

M365 관리센터 내보내기는 값을 **이중 따옴표로 한 번 더 감싸고**(`"""TestApp"""`),
`keyCredentials` 같은 필드에 **JSON을 통째로** 넣는다.

브라우저에 CSV 파서가 없으므로 RFC4180 파서를 직접 구현하고,
파싱 후 남는 리터럴 따옴표를 벗기되 JSON 시작(`[`/`{`)은 건드리지 않도록 처리했다.

### 3-6. 실제 테넌트 데이터 유출 방지

Tab 3 테스트에 실제 앱 등록 CSV를 사용했다. 저장소가 public이므로:

- 테스트 후 복사본 즉시 삭제
- `.gitignore`에 `AppRegistrationList*.csv`, `entra_export/`, `0N_*.csv`, `.testdata/` 등록
- 커밋 전 `git status`로 미포함 확인

---

## 4. 앱 등록 CSV 실사 결과 (499건)

### 4-1. 이 CSV로는 못 하는 것

| 빠진 필드 | 못 하는 것 |
|---|---|
| `requiredResourceAccess` | **Graph API 권한 유무 판정** ← "깡통 App" 정의의 핵심 |
| `owners` | **소유자 식별** ← 확인 메일 발송 대상 |
| `signInActivity` | 마지막 사용 시각 → 미사용 판정 |

> 요청한 권한(`requiredResourceAccess`) ≠ 실제 부여된 권한(`appRoleAssignment`).
> 회수 판단은 실제 부여된 권한 기준이어야 하므로 둘 다 수집해 대조한다.

### 4-2. 그래도 나온 것

| 항목 | 건수 |
|---|---:|
| 자격증명 없음 (인증서·시크릿 전무) | **372 (75%)** |
| 전부 만료 | 24 |
| 90일 내 만료 | 17 |
| 유효 자격증명 보유 | **86 (17%)** |
| 자격증명 + 리디렉션 URI 모두 없음 | **351 (70%)** |
| 암묵적 허용 흐름 활성화 | 49 |
| 멀티테넌트 / 개인계정 허용 | 41 / 10 |

**자격증명이 없으면 토큰 발급이 물리적으로 불가능하다.** 권한 목록을 몰라도 372건은
"현재 사용 불가"가 확정이다. 실제로 살아 있는 앱은 **86건**뿐이다.

### 4-3. 가장 중요한 발견

| tag | 건수 |
|---|---:|
| `AIAgentBuilder` / `AgentCreatedBy:CopilotStudio` / `AgenticApp` | **305 (61%)** |
| `Microsoft Fabric Identity` | 15 |

**전체의 61%가 Copilot Studio 자동 생성 앱이다.**

기존 PPT는 미사용 앱 312개의 원인을 *"대부분 교육용 일시 계정"* 으로 기록했으나,
**305 ≈ 312** 로 숫자가 거의 일치한다. 실제 원인은 교육용 계정이 아니라
**Copilot Studio가 봇 생성 시 앱 등록을 자동 생성하고, 봇을 삭제해도 앱 등록은 남는 것**일 가능성이 높다.

원인이 다르면 대책도 다르다.

- 교육용이 원인 → 사전 안내 강화
- **Copilot Studio가 원인** → 개인 소유자에게 안내 메일이 무의미.
  환경 단위 정리 정책, 봇 삭제 시 앱 등록 동반 삭제, 생성 권한 통제가 필요

Tab 3은 이를 반영해 **자동 생성 앱을 개별 메일 대상에서 자동 제외**하고 공지 방식을 권고한다.

생성 추이도 이를 뒷받침한다 — 2022년 3건 / 2023년 4건 / 2024년 55건 / **2025년 313건** / 2026년(8월) 124건.
Copilot Studio 도입 시점과 겹친다. **유입을 막지 않으면 정리해도 다시 쌓인다.**

---

## 5. 검증 이력

모든 검증은 실제 브라우저에서 전 구간을 실행해 수행했다.

### Tab 1 — 17개 케이스

| TC | 검증 내용 | 결과 |
|---|---|---|
| TC-1 | `Mail.Read`/Application/"특정 사서함" → 재판정 → `SDR` **G1** + SEC-SCOPE | ✅ |
| TC-2 | 전사 메일 수집 → `SAR` **G3**, RISK-MASSREAD Critical | ✅ |
| TC-5 | 본문 제외 + 미저장 → `SAW` G3 → **G2** (1등급 하향 캡) | ✅ |
| TC-6 | 임원 Daily 발송 → `GAR` G2 → **G3** | ✅ |
| TC-8 | `Directory.ReadWrite.All` + 퇴직자 제거 → **G3** + Critical 2건 | ✅ |
| TC-14 | "특정 팀" + "전사 모든 팀" → 재판정 미발동 + 범위 상충 | ✅ |
| TC-12 | DRM 암호화 PPTX → 파싱 전 차단 + 안내 | ✅ |
| TC-16 | UI 편집 → 즉시 재판정 (G1 → G3 전환) | ✅ |
| TC-17 | `DecompressionStream` 제거 후 자체 inflate → **바이트 단위 일치** | ✅ |
| 그 외 | TC-3·4·7·9·10·11·13·15·18·19 | ✅ |

### Tab 2

136종 로드, 범주·등급 필터, 한글/영문 검색 랭킹, 셀렉트박스, 표 클릭 → 상세.
사용자 제시 5건 검증 — Power Platform `SWP`→A3 · Global Reader `TRN`→A2 ·
AI Administrator `TWP`→A3 · Compliance Administrator `TWP`→A3 · Security Administrator `TWP`→A3.

### Tab 3

실제 관리센터 CSV 499건 처리 — D1 321 / D2 75 / D4 103, 자동생성 323건, 보안점검 115건.
샘플 데이터로 D3 판정·소유자별 메일·필터·CSV/JSON 내보내기 확인.

### 공통

콘솔 에러 0건 · 데스크톱(1280px)·모바일(375px) 가로 넘침 0건 · 탭 전환 시 상태 유지 ·
배포 URL에서 동일 동작 확인.

### PPT

PowerPoint COM으로 14장 전체를 PNG 렌더링해 육안 검수. 넘침 2건 수정 후 재검수 완료.

---

## 6. 산출물

### 도구

| 경로 | 설명 |
|---|---|
| `GraphAPI_등급판정시스템.html` | 본체 (2,784줄, 3탭, 단일 파일·오프라인) |
| `index.html` | 웹 배포 진입점 |

### 문서

| 경로 | 설명 |
|---|---|
| `docs/BRD_M365_권한관리체계.md` | 배경·문제정의·목표·범위·성공지표·리스크 |
| `docs/PRD_GAPS_통합판정시스템.md` | 기능요구사항·판정 규칙·데이터 모델·화면·테스트 |
| `docs/HISTORY.md` | 본 문서 |
| `docs/모니터링_데이터수집_가이드.md` | CSV 실측 + 수집 방법 + Graph 활동 로그 + KQL |
| `docs/M365 관리자 권한 분류 및 등급 기준_v1.0.pptx` | 기준 장표 14장 |
| `docs/M365 관리자 역할 명세서_v1.0.xlsx` | 역할 명세 136건 3시트 |
| `README.md` | 사용법·기준 요약·수정 방법 |

### 템플릿 · 샘플

| 경로 | 설명 |
|---|---|
| `templates/Graph API 권한 신청서_템플릿_v1.0.pptx` | 요청자 배포용 빈 양식 (Full + 약식) |
| `samples/*.pptx` | 작성 예시 6종 (테스트 케이스 겸용) |

### 스크립트

| 경로 | 설명 |
|---|---|
| `tools/admin_roles_data.py` | 관리자 역할 데이터 **단일 원본** (136종) |
| `tools/build_admin_role_docs.py` | PPT·Excel·JSON 생성 + HTML 주입 |
| `tools/build_templates.py` | 신청 템플릿·샘플 PPTX 생성 |
| `tools/collect_entra_inventory.ps1` | Entra 인벤토리 수집 (CSV 5종) |

---

## 7. 남은 과제

| 순서 | 작업 | 비고 |
|---|---|---|
| **0** | **Graph 활동 로그 진단 설정 ON** | ⚠ **소급 불가** — 가장 먼저. Entra ID P1 + Log Analytics |
| 1 | `collect_entra_inventory.ps1` 실행 | 모듈 설치·로그인이 대화형이라 담당자 직접 실행 필요 |
| 2 | Copilot Studio 자동 생성 앱 305건 성격 확인 | 개별 안내 대상인지 일괄 정리 대상인지 판단 |
| 3 | 소유자 없는 앱 처리 방안 | 감사 로그로 생성자 추적 또는 담당 조직 확인 |
| 4 | 정보보호그룹 민감정보 기준 확정 → `CFG.DICT` 반영 | PRD Phase 2 |
| 5 | Vercel GitHub 연동 승인 → 푸시 시 자동 배포 | 현재는 CLI 수동 배포 |
| 6 | IR 시스템 연동, 요청 대장 자동 적재 | PRD Phase 3 |

### 검토만 하고 만들지 않은 것

Tab 3 주제를 정할 때 함께 검토한 후보들. 향후 참고용으로 남긴다.

| 후보 | 성격 | 비고 |
|---|---|---|
| 최소권한 역제안기 | 업무 → 필요 권한 (사전 예방) | MS `delegate-by-task` 문서가 근거 |
| 권한 조합 위험 분석 | 개별로는 안전하나 합치면 위험한 조합 (SoD) | 현 체계 사각지대 |
| 침해 영향도 시뮬레이터 | 유출 시 영향 서술 (설득 자료) | Committee 상정용 |

---

## 8. 참고 문서

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

**모니터링**
- [Microsoft Graph 활동 로그](https://learn.microsoft.com/ko-kr/entra/identity/monitoring-health/concept-microsoft-graph-activity-logs)
- [Entra 진단 설정 구성](https://learn.microsoft.com/ko-kr/entra/identity/monitoring-health/howto-configure-diagnostic-settings)

**사내**
- `Graph API 할당 프로세스 및 기준_260806_학습용.pptx`
- `M365관련 요청 사항 정리_260804.xlsx`
- `AppRegistrationList (1).csv` (2026-08-06 내보내기, 499건)
