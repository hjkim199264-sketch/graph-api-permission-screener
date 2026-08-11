# Graph API 권한 요청 대응 체계 — Day2 산출물

DX혁신실 M365 운영 | 2026-08-11

`Graph API 할당 프로세스 및 기준_260806_학습용.pptx` 의 등급 분류 기준을 문서화(BRD/PRD)하고,
신청 템플릿 PPTX를 올리면 **등급 산정 · 보완 요청사항 · 향후 진행사항** 을 자동 산출하는 도구를 구현한 결과물이다.

---

## 바로 쓰기

1. **`GraphAPI_등급판정시스템.html`** 을 더블클릭 (Edge/Chrome)
2. `samples/` 안의 PPTX를 창에 끌어다 놓기 — 또는 **[샘플로 체험하기]** 클릭
3. 결과 확인 후 **[IR 회신문 복사]** 로 그대로 회신

> 서버·설치·인터넷 연결 불필요. 업로드한 파일은 **브라우저 안에서만** 처리되며 외부로 전송되지 않는다.

---

## 파일 구성

| 경로 | 설명 |
|---|---|
| `GraphAPI_등급판정시스템.html` | **판정 시스템 본체.** 단일 파일·오프라인 동작 |
| `docs/BRD_GraphAPI_권한관리체계.md` | 배경·문제정의·비즈니스 목표·범위·성공지표·리스크 |
| `docs/PRD_GraphAPI_등급판정시스템.md` | 기능요구사항·분류 규칙·API 사전·데이터 모델·화면정의·테스트 결과 |
| `templates/Graph API 권한 신청서_템플릿_v1.0.pptx` | **요청자 배포용 빈 양식** (Full 1p + 약식 1p) |
| `samples/*.pptx` | 작성 예시 6종 (테스트 케이스 겸용) |
| `tools/build_templates.py` | 위 템플릿·샘플 PPTX 생성 스크립트 |
| `.claude/launch.json` | 로컬 테스트용 정적 서버 설정 (운영에 불필요) |

---

## 판정 로직 요약

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

HTML 파일 상단의 `const CFG = { … }` **설정 블록만** 수정하면 된다. 그 외 코드는 손댈 필요가 없다.

| 항목 | 위치 | 용도 |
|---|---|---|
| `CFG.DICT` | API 사전 | 권한별 정보유형(G/S/GS). 슬라이드 7의 23건이 정본 |
| `CFG.FAMILY` | 패밀리 규칙 | 사전 미등재 권한의 정보유형 추정 |
| `CFG.KW_LIMIT` / `KW_TENANT` | 재판정 키워드 | 대상 한정 표현 / 전사 표현 |
| `CFG.RE_*` | 2차 분류 판정 정규식 | 최소화·동의·저장·규모 |
| `CFG.PRIVESC` / `IMPERSONATE` / `MASSREAD` | 고위험 권한 목록 | Critical 판정 |
| `CFG.GRADE_INFO` / `CFG.STEPS` | 등급별 절차 | 담당 조직·소요·단계 |

정보보호그룹의 민감정보 기준이 확정되면 `CFG.DICT` 의 `info` 값을 교체하면 된다.

---

## 알아둘 점

- **판정 결과는 초안(Draft)** 이다. 최종 승인은 정보보호그룹 검토와 IR 결재선을 따른다.
- 파싱이 틀리면 화면 **② 추출된 신청 정보** 에서 직접 고칠 수 있고, 수정 즉시 재판정된다.
- **DRM/IRM(민감도 레이블) 암호화 PPTX는 읽을 수 없다.** PowerPoint에서 열어 레이블을 해제한 뒤 올려야 한다.
- 레거시 `.ppt` 는 미지원 — `.pptx` 로 저장 후 사용.
- 사전에 없는 권한은 **보수적으로 민감(S)** 으로 가정하고 확인을 요청한다.

## 참고

- [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Limiting application permissions to specific mailboxes](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
- [Sites.Selected 개요](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)
