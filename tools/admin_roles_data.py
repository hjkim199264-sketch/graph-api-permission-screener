# -*- coding: utf-8 -*-
"""
M365 / Microsoft Entra 관리자 역할 등급 판정 데이터셋 (단일 원본)

이 파일이 PPT · Excel · HTML(Tab2) 3종 산출물의 공통 원본이다.
기준을 바꾸려면 이 파일만 수정한 뒤 build_admin_role_docs.py 를 다시 실행한다.

출처
  - Microsoft 365 관리 센터 관리자 역할 정보
    https://learn.microsoft.com/ko-kr/microsoft-365/admin/add-users/about-admin-roles
  - Microsoft Entra 기본 제공 역할 (109종, 권한 있는 역할 표시 포함)
    https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/permissions-reference
  - 권한 있는 역할 및 사용 권한
    https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/privileged-roles-permissions

────────────────────────────────────────────────────────────────────────────
등급 판정 기준 — 3기준 × 2값 = 8조합 → A1 / A2 / A3
(사내 Graph API 등급 체계 G1/G2/G3와 동일한 구조로 설계)

  ① 영향 범위 (Scope)
      S = 서비스·기능 한정   단일 제품 또는 단일 기능 영역에만 적용
      T = 테넌트 전역        디렉터리 전반(사용자·그룹·역할·앱·정책·도메인) 또는 전 서비스 교차

  ② 작업 유형 (Operation)
      R = 읽기 전용          조회·모니터링·보고만 가능
      W = 구성 변경          생성·수정·삭제·정책 변경 가능

  ③ 민감 자산 접근 (Sensitive)  ← 두 경로 중 하나라도 해당하면 P
      N = 없음
      P = 있음
          (a) 권한 상승 경로 : 역할 할당, 자격 증명·MFA 변경, 앱 권한 동의,
                               조건부 액세스·인증 정책 변경, 도메인·페더레이션 변경
          (b) 콘텐츠 접근 경로 : 메일·파일·채팅·문서 원문을 열람하거나
                               열람 가능한 위임·사이트 관리자 지정이 가능

  매핑
      민감 자산 접근 P 있음            →  A3     (범위·작업과 무관)
      N + 읽기(R) + 서비스(S)          →  A1
      N + 읽기(R) + 전역(T)            →  A2
      N + 변경(W) + 서비스(S)          →  A2
      N + 변경(W) + 전역(T)            →  A3

  핵심 : Graph API 기준에서 "Delegated면 무조건 G1"이 한 축으로 등급을 결정했듯,
        관리자 역할에서는 "민감 자산 접근이 있으면 무조건 A3"가 등급을 결정한다.
        권한 상승과 콘텐츠 열람은 읽기/쓰기·범위와 무관하게 최고 위험이기 때문이다.
────────────────────────────────────────────────────────────────────────────
"""

# 범주 (Microsoft 365 관리 센터 '범주별로 모두 표시' 기준)
CATEGORIES = ["전역", "ID", "공동 작업", "보안 및 준수", "디바이스", "읽기 전용", "청구", "기타"]

GRADE_DEF = {
    "A1": dict(risk="🟢 낮음", label="제한적 조회", color="1B7F3B", bg="E7F5EC",
               desc="단일 기능 영역의 읽기 전용. 조직 데이터·설정을 변경할 수 없다.",
               process="운영그룹 자체 승인",
               review="보안 검토 생략 (사후 이력 관리만)",
               pim="상시 할당 가능", eta="1영업일"),
    "A2": dict(risk="🟡 중간", label="서비스 운영 / 전역 조회", color="B26A00", bg="FEF3E2",
               desc="단일 서비스의 구성을 변경하거나, 테넌트 전반을 조회할 수 있다. "
                    "권한 상승·콘텐츠 열람 경로는 없다.",
               process="운영그룹 주관 + 정보보호그룹 공유",
               review="보안 검토 공유 (의견 수렴)",
               pim="PIM 등록 권장", eta="3~5영업일"),
    "A3": dict(risk="🔴 높음", label="특권 / 전역 변경", color="B3261E", bg="FDECEA",
               desc="권한 상승 경로 또는 사용자 콘텐츠 접근 경로를 가지거나, "
                    "테넌트 전역 구성을 변경할 수 있다.",
               process="정보보호그룹 보안 검토 필수, 필요시 Committee 상정",
               review="보안 검토 필수 (승인 없이는 부여 불가)",
               pim="PIM 필수 · 상시 할당 금지 · MFA 필수", eta="2주 이상"),
}

# 민감 자산 접근(P) 판정 사유 코드
SENS_REASON = {
    "ESC-ROLE":  "역할 할당·PIM 관리 가능 (권한 상승)",
    "ESC-CRED":  "타 계정의 암호·MFA 등 인증 정보 변경 가능 (계정 탈취 경로)",
    "ESC-APP":   "앱 등록·앱 권한 동의 부여 가능 (권한 상승)",
    "ESC-POLICY":"조건부 액세스·인증 정책 변경 가능 (보안 통제 무력화)",
    "ESC-DOMAIN":"도메인·페더레이션 설정 변경 가능 (신뢰 체계 훼손)",
    "ESC-DEVICE":"디바이스에 스크립트·구성 배포 가능 (엔드포인트 장악)",
    "CONTENT":   "메일·파일·채팅·문서 원문 접근 또는 접근 위임 설정 가능",
    "CONTENT-EDISC":"eDiscovery·콘텐츠 검색으로 전사 데이터 열람 가능",
}

# ─────────────────────────────────────────────────────────────────────────────
# 역할 데이터
#   (영문명, 한글명, 범주, 범위 S/T, 작업 R/W, 민감 N/P, MS공식_권한있는역할, 민감사유코드, 핵심 리스크)
# ─────────────────────────────────────────────────────────────────────────────
ROLES = [
# ── 전역 / ID ────────────────────────────────────────────────────────────────
("Global Administrator","전역 관리자","전역","T","W","P",True,"ESC-ROLE",
 "테넌트 전체에 무제한 권한. 모든 사용자·관리자 암호 재설정, 모든 설정 변경, Azure 구독 권한 상승까지 가능한 최고 위험 역할"),
("Privileged Role Administrator","권한 있는 역할 관리자","ID","T","W","P",True,"ESC-ROLE",
 "모든 Entra 역할 할당과 PIM 전체를 관리. 스스로 전역 관리자를 부여할 수 있어 사실상 전역 관리자와 동급"),
("Privileged Authentication Administrator","권한 있는 인증 관리자","ID","T","W","P",True,"ESC-CRED",
 "전역 관리자를 포함한 모든 사용자의 인증 방법을 재설정 가능. 관리자 계정 탈취의 직접 경로"),
("User Administrator","사용자 관리자","ID","T","W","P",True,"ESC-CRED",
 "사용자 계정 생성·비활성화, 비관리자 암호 재설정, 그룹 관리. 계정 탈취 및 접근 권한 확대 경로"),
("Authentication Administrator","인증 관리자","ID","T","W","P",True,"ESC-CRED",
 "비관리자 사용자의 MFA·인증 방법을 조회·재설정 가능. MFA 우회를 통한 계정 탈취 경로"),
("Authentication Policy Administrator","인증 정책 관리자","ID","T","W","P",True,"ESC-POLICY",
 "MFA 설정·인증 방법 정책·암호 보호 정책 변경 가능. 조직 전체 인증 통제를 약화시킬 수 있음"),
("Authentication Extensibility Administrator","인증 확장성 관리자","ID","T","W","P",True,"ESC-POLICY",
 "사용자 지정 인증 확장을 생성·관리. 외부 시스템으로 인증 흐름을 우회시킬 수 있음"),
("Authentication Extensibility Password Administrator","인증 확장성 암호 관리자","ID","T","W","P",True,"ESC-CRED",
 "사용자 지정 인증의 암호 제출 이벤트를 트리거 가능"),
("Password Administrator","암호 관리자","ID","T","W","P",True,"ESC-CRED",
 "비관리자 및 암호 관리자의 암호를 재설정 가능. 계정 탈취 경로"),
("Helpdesk Administrator","기술 지원팀 관리자","ID","T","W","P",True,"ESC-CRED",
 "비관리자 암호 재설정 및 강제 로그아웃. 지원 업무 명목의 광범위한 계정 접근 경로"),
("Conditional Access Administrator","조건부 액세스 관리자","ID","T","W","P",True,"ESC-POLICY",
 "조건부 액세스 정책 전체를 관리. MFA·디바이스 준수 요구를 해제해 전사 보안 통제를 무력화할 수 있음"),
("Application Administrator","애플리케이션 관리자","ID","T","W","P",True,"ESC-APP",
 "앱 등록·엔터프라이즈 앱 전체 관리 및 위임/애플리케이션 권한 동의 부여 가능 (Graph 앱 권한 제외). 권한 상승 경로"),
("Cloud Application Administrator","클라우드 애플리케이션 관리자","ID","T","W","P",True,"ESC-APP",
 "앱 프록시를 제외한 앱 등록·엔터프라이즈 앱 관리 및 권한 동의 부여 가능"),
("Application Developer","애플리케이션 개발자","ID","T","W","P",True,"ESC-APP",
 "테넌트 설정과 무관하게 앱 등록 생성 가능. 무단 앱을 통한 데이터 반출 경로"),
("Hybrid Identity Administrator","하이브리드 ID 관리자","ID","T","W","P",True,"ESC-DOMAIN",
 "AD-Entra 클라우드 프로비저닝 및 페더레이션 설정 관리. 온프레미스 신뢰 체계 훼손 가능"),
("Domain Name Administrator","도메인 이름 관리자","ID","T","W","P",True,"ESC-DOMAIN",
 "클라우드·온프레미스 도메인 관리. 도메인 검증을 통한 신뢰 탈취 경로"),
("External Identity Provider Administrator","외부 ID 공급자 관리자","ID","T","W","P",True,"ESC-DOMAIN",
 "직접 페더레이션용 ID 공급자 구성. 외부 IdP를 통한 우회 로그인 경로"),
("Identity Governance Administrator","Identity Governance 관리자","ID","T","W","P",True,"ESC-ROLE",
 "액세스 패키지·액세스 검토 등 ID 거버넌스 전반 관리. 접근 권한 부여 흐름을 좌우"),
("Lifecycle Workflow Administrator","수명 주기 워크플로 관리자","ID","T","W","P",True,"ESC-ROLE",
 "입·퇴사 자동화 워크플로 관리. 계정 생성·권한 부여를 자동 실행할 수 있음"),
("Tenant Governance Administrator","테넌트 거버넌스 관리자","ID","T","W","P",True,"ESC-ROLE",
 "테넌트 거버넌스 서비스 전체 기능 관리"),
("Directory Writer","디렉터리 작성기","ID","T","W","P",True,"ESC-ROLE",
 "기본 디렉터리 정보 읽기·쓰기. 서비스 계정 용도이며 사용자 직접 할당 비권장"),
("Directory Synchronization Account","디렉터리 동기화 계정","ID","T","W","P",True,"ESC-ROLE",
 "Entra Connect 서비스 전용. 사람에게 절대 할당하지 말 것"),
("Partner Tier1 Support","파트너 계층1 지원","ID","T","W","P",True,"ESC-CRED",
 "Microsoft 내부 전용. 일반 사용 금지 (Do not use)"),
("Partner Tier2 Support","파트너 계층2 지원","ID","T","W","P",True,"ESC-CRED",
 "Microsoft 내부 전용. 일반 사용 금지 (Do not use)"),
("Agent Identity Administrator","에이전트 ID 관리자","ID","T","W","P",True,"ESC-ROLE",
 "에이전트 ID·청사진·보안 주체 전체 관리. AI 에이전트에 권한을 부여하는 경로"),
("Agent Identity Developer","에이전트 ID 개발자","ID","T","W","P",False,"ESC-APP",
 "에이전트 ID 청사진 및 보안 주체 생성 가능"),
("Attribute Provisioning Administrator","특성 프로비저닝 관리자","ID","T","W","P",True,"ESC-APP",
 "앱의 사용자 지정 보안 특성 프로비저닝 구성을 변경 가능"),
("Tenant Creator","테넌트 작성자","ID","T","W","N",False,"",
 "신규 Entra / B2C 테넌트 생성 가능. 관리 사각지대 테넌트 생성 위험"),
("Groups Administrator","그룹 관리자","ID","T","W","N",False,"",
 "모든 그룹 생성·편집·삭제 및 그룹 정책 관리. 그룹 멤버십 변경으로 접근 권한이 간접 확대될 수 있음"),
("Guest Inviter","게스트 초대자","ID","S","W","N",False,"",
 "외부 게스트 사용자 초대 가능. 외부인 유입 통제 필요"),
("Attribute Assignment Administrator","특성 할당 관리자","ID","T","W","N",False,"",
 "개체에 사용자 지정 보안 특성 할당. 특성 기반 접근 정책에 영향"),
("Attribute Definition Administrator","특성 정의 관리자","ID","T","W","N",False,"",
 "사용자 지정 보안 특성의 정의를 관리"),
("People Administrator","People 관리자","ID","S","W","N",False,"",
 "전 사용자의 프로필 사진·대명사·이름 발음·프로필 카드 설정 변경"),
("People Manager","사용자 프로필 관리자 (People Manager)","ID","S","W","N",False,"",
 "조직 전반의 사용자 프로필 관리"),
("Permissions Management Administrator","권한 관리 관리자","ID","T","W","N",False,"",
 "Entra 자격 관리(entitlement management) 전반 관리"),
("Extended Directory User Administrator","확장 디렉터리 사용자 관리자","ID","S","W","N",False,"",
 "Teams 확장 디렉터리의 외부 사용자 프로필 관리"),
("External ID User Flow Administrator","외부 ID 사용자 흐름 관리자","ID","S","W","N",False,"",
 "외부 ID 사용자 흐름 생성·관리"),
("External ID User Flow Attribute Administrator","외부 ID 사용자 흐름 특성 관리자","ID","S","W","N",False,"",
 "사용자 흐름에서 사용하는 특성 스키마 관리"),
("B2C IEF Keyset Administrator","B2C IEF 키 세트 관리자","ID","S","W","P",True,"ESC-CRED",
 "Identity Experience Framework의 페더레이션·암호화 비밀 관리"),
("B2C IEF Policy Administrator","B2C IEF 정책 관리자","ID","S","W","N",False,"",
 "Identity Experience Framework 트러스트 프레임워크 정책 관리"),

# ── 보안 및 준수 ─────────────────────────────────────────────────────────────
("Security Administrator","보안 관리자","보안 및 준수","T","W","P",True,"ESC-POLICY",
 "Defender·Entra ID 보호·Purview 등 보안 구성 전반을 변경. 보안 정책·경보 설정을 무력화할 수 있는 특권 역할"),
("Security Operator","보안 연산자","보안 및 준수","T","W","P",True,"ESC-CRED",
 "보안 이벤트 관리 및 ID 격리 조치(계정 차단·세션 해지) 수행 가능"),
("Security Reader","보안 읽기 권한자","읽기 전용","T","R","N",True,"",
 "Defender·ID 보호·PIM·로그인/감사 로그·Purview 보안 정보 조회. 읽기 전용이나 보안 태세 전반이 노출됨"),
("Compliance Administrator","준수 관리자","보안 및 준수","T","W","P",False,"CONTENT-EDISC",
 "Purview 규정 준수 구성·보고 관리. DLP·보존·eDiscovery 설정을 통해 전사 콘텐츠에 접근할 수 있는 경로"),
("Compliance Data Administrator","규정 준수 데이터 관리자","보안 및 준수","T","W","P",False,"CONTENT-EDISC",
 "규정 준수 콘텐츠 생성·관리. 전사 데이터 검색·보존 대상 지정 가능"),
("Purview Workload Content Manager","Purview 워크로드 콘텐츠 관리자","보안 및 준수","S","W","P",False,"CONTENT",
 "Purview 포털에서 M365 데이터를 관리하거나 제거 가능"),
("Purview Workload Content Writer","Purview 워크로드 콘텐츠 기록기","보안 및 준수","S","W","P",False,"CONTENT",
 "Purview 포털에서 M365 데이터 조회 및 편집 가능"),
("Purview Workload Content Reader","Purview 워크로드 콘텐츠 읽기 권한자","보안 및 준수","S","R","P",False,"CONTENT",
 "Purview 포털에서 M365 데이터 조회 가능. 읽기 전용이나 콘텐츠 원문에 접근"),
("Azure Information Protection Administrator","Azure Information Protection 관리자","보안 및 준수","S","W","N",False,"",
 "민감도 레이블·보호 정책 관리. 보호 정책 변경으로 문서 암호화가 해제될 수 있어 상향 검토 필요"),
("Cloud App Security Administrator","Cloud App Security 관리자","보안 및 준수","S","W","N",False,"",
 "Defender for Cloud Apps 전반 관리. 세션 정책·앱 차단 설정 변경"),
("Attack Simulation Administrator","공격 시뮬레이션 관리자","보안 및 준수","S","W","N",False,"",
 "피싱 시뮬레이션 캠페인 생성·관리. 실제 사용자 대상 메일 발송이 수반됨"),
("Attack Payload Author","공격 페이로드 작성자","보안 및 준수","S","W","N",False,"",
 "공격 시뮬레이션 페이로드 작성"),
("Customer Lockbox Access Approver","고객 Lockbox 액세스 승인자","보안 및 준수","T","W","P",False,"CONTENT",
 "Microsoft 지원의 고객 데이터 접근 요청을 승인. 승인 시 외부인이 조직 데이터에 접근"),
("Global Secure Access Administrator","전역 보안 액세스 관리자","보안 및 준수","T","W","P",False,"ESC-POLICY",
 "Global Secure Internet/Private Access 전반 관리. 네트워크 접근 경로 통제"),
("Global Secure Access Log Reader","전역 보안 액세스 로그 판독기","읽기 전용","S","R","N",False,"",
 "네트워크 트래픽 로그 읽기 전용 접근"),
("Attribute Log Administrator","특성 로그 관리자","보안 및 준수","S","W","N",False,"",
 "사용자 지정 보안 특성 관련 감사 로그 조회 및 진단 설정 구성"),
("Attribute Log Reader","특성 로그 판독기","읽기 전용","S","R","N",False,"",
 "사용자 지정 보안 특성 관련 감사 로그 조회"),

# ── 공동 작업 (워크로드) ─────────────────────────────────────────────────────
("Exchange Administrator","Exchange 관리자","공동 작업","S","W","P",False,"CONTENT",
 "전 사용자 사서함 관리. 삭제 항목 복구, '다른 이름으로 보내기'·'대신 보내기' 위임 설정으로 타인 메일에 접근 가능"),
("Exchange Recipient Administrator","Exchange 받는 사람 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Exchange Online 받는 사람 생성·수정. 사서함 속성 변경을 통한 접근 경로"),
("SharePoint Administrator","SharePoint 관리자","공동 작업","S","W","P",False,"CONTENT",
 "전 사이트·사이트 모음 관리. 자신을 사이트 관리자로 지정해 모든 문서에 접근 가능"),
("SharePoint Advanced Management Administrator","SharePoint 고급 관리 관리자","공동 작업","S","W","P",False,"CONTENT",
 "SharePoint 고급 관리 전반. 파일·폴더 이름/경로/URL 조회 및 권한 제거 가능"),
("SharePoint Embedded Administrator","SharePoint Embedded 관리자","공동 작업","S","W","N",False,"",
 "SharePoint Embedded 컨테이너 관리"),
("Teams Administrator","Teams 관리자","공동 작업","S","W","N",False,"",
 "Teams 서비스 전반 관리. 모임·페더레이션·조직 전체 설정 변경"),
("Teams Communications Administrator","Teams 통신 관리자","공동 작업","S","W","N",False,"",
 "Teams 통화·모임 기능 관리"),
("Teams Telephony Administrator","Teams 전화 통신 관리자","공동 작업","S","W","N",False,"",
 "Teams 음성·전화 기능 관리 및 문제 해결"),
("Teams Device Administrator","Teams 디바이스 관리자","디바이스","S","W","N",False,"",
 "Teams 인증 디바이스 관리 작업 수행"),
("Teams External Collaborations Administrator","Teams 외부 공동 작업 관리자","공동 작업","S","W","N",False,"",
 "Teams 외부 공동 작업 정책·설정 관리. 외부 조직과의 통신 허용 범위 결정"),
("Teams Communications Support Engineer","Teams 통신 지원 엔지니어","공동 작업","S","R","N",False,"",
 "고급 도구로 Teams 통신 문제 해결. 통화 상세 기록 조회"),
("Teams Communications Support Specialist","Teams 통신 지원 전문가","공동 작업","S","R","N",False,"",
 "기본 도구로 Teams 통신 문제 해결"),
("Teams Reader","Teams 읽기 권한자","읽기 전용","S","R","N",False,"",
 "Teams 관리 센터 전체 조회 (변경 불가)"),
("AI Administrator","AI 관리자","공동 작업","T","W","P",True,"ESC-APP",
 "Copilot·AI 에이전트 전반 관리. 앱·에이전트에 테넌트 전체 동의를 부여할 수 있어 권한 상승 경로 (Graph 앱 권한은 제외)"),
("AI Reader","AI 리더","읽기 전용","T","R","N",True,"",
 "Copilot·AI 에이전트 전반 및 에이전트 ID 조회. 읽기 전용"),
("Power Platform Administrator","Power Platform 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Power Apps·Power Automate·Power BI·Fabric 및 DLP 전체 관리. 환경 관리 권한을 통해 Dataverse 데이터에 접근 가능"),
("Fabric Administrator","패브릭 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Fabric·Power BI 전체 관리 기능 및 감사 검토. 작업 영역 데이터 접근 경로"),
("Dynamics 365 Administrator","Dynamics 365 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Dynamics 365 Online 전체 관리. 업무 데이터 접근 경로"),
("Dynamics 365 Business Central Administrator","Dynamics 365 Business Central 관리자","기타","S","W","P",False,"CONTENT",
 "Business Central 환경의 모든 관리 작업 수행. 재무·업무 데이터 접근 경로"),
("Yammer Administrator","Yammer 관리자","공동 작업","S","W","N",False,"",
 "Viva Engage(Yammer) 서비스 전반 관리"),
("Skype for Business Administrator","비즈니스용 Skype 관리자","공동 작업","S","W","N",False,"",
 "비즈니스용 Skype 전반 관리"),
("Kaizala Administrator","Kaizala 관리자","기타","S","W","N",False,"",
 "Microsoft Kaizala 설정 관리"),
("Office Apps Administrator","Office 앱 관리자","공동 작업","S","W","N",False,"",
 "Office 클라우드 정책 관리, Office 스크립트 설정 관리"),
("Search Administrator","검색 관리자","공동 작업","S","W","N",False,"",
 "Microsoft Search 구성 및 검색 결과 콘텐츠 전체 관리"),
("Search Editor","검색 편집기","공동 작업","S","W","N",False,"",
 "북마크·Q&A 등 편집 콘텐츠 생성·관리"),
("Knowledge Admin","지식 관리자 (Knowledge Admin)","공동 작업","S","W","N",False,"",
 "지식·학습 등 인텔리전트 기능 구성"),
("Knowledge Manager","지식 매니저 (Knowledge Manager)","공동 작업","S","W","N",False,"",
 "토픽·지식 생성·관리·승격"),
("Microsoft 365 Migration Administrator","마이그레이션 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Google Drive·Box·Dropbox·Slack 등에서 M365로 콘텐츠 마이그레이션. 이관 콘텐츠 접근 경로"),
("Microsoft Graph Data Connect Administrator","Microsoft Graph Data Connect 관리자","공동 작업","T","W","P",False,"CONTENT",
 "Graph Data Connect 설정 및 앱 권한 부여 요청 승인. 대량 M365 데이터를 Azure로 반출하는 경로"),
("Microsoft 365 Backup Administrator","Microsoft 365 백업 관리자","공동 작업","S","W","P",False,"CONTENT",
 "M365 백업·복원 관리. 백업본을 통한 콘텐츠 접근 경로"),
("Exchange Backup Administrator","Exchange 백업 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Exchange 콘텐츠 백업 및 세분화된 복원. 메일 콘텐츠 접근 경로"),
("SharePoint Backup Administrator","SharePoint 백업 관리자","공동 작업","S","W","P",False,"CONTENT",
 "SharePoint·OneDrive 백업 및 복원. 문서 콘텐츠 접근 경로"),
("Entra Backup Administrator","Entra Backup 관리자","ID","S","W","N",False,"",
 "Microsoft Entra 백업 전반 관리"),
("Entra Backup Reader","Entra Backup 판독기","읽기 전용","S","R","N",False,"",
 "Microsoft Entra 백업 조회"),
("Organizational Message Approver","조직 메시지 승인자","공동 작업","S","W","N",False,"",
 "전 직원 대상 조직 메시지 검토·승인·거부"),
("Organizational Message Writer","조직 메시지 작성자","공동 작업","S","W","N",False,"",
 "Microsoft 제품 화면에 표시되는 조직 메시지 작성·게시"),
("Organizational Data Source Administrator","조직 데이터 원본 관리자","공동 작업","S","W","N",False,"",
 "M365로 수집되는 조직 데이터 원본 설정·관리"),
("Organization Branding Administrator","조직 브랜딩 관리자","기타","S","W","N",False,"",
 "테넌트 로그인 화면 등 조직 브랜딩 관리. 피싱 유사 화면 구성 우려"),
("Places Manager","위치 관리자","기타","S","W","N",False,"",
 "Microsoft Places 서비스 전반 관리"),
("Virtual Visits Administrator","Virtual Visits 관리자","공동 작업","S","W","N",False,"",
 "Virtual Visits 정보·메트릭 관리 및 공유"),
("Viva Glint Tenant Administrator","Viva Glint 테넌트 관리자","공동 작업","S","W","P",False,"CONTENT",
 "Viva Glint 설정 관리. 임직원 설문·피드백 데이터 접근 경로"),
("Viva Goals Administrator","Viva Goals 관리자","공동 작업","S","W","N",False,"",
 "Viva Goals 전반 및 구성 관리"),
("Viva Pulse Administrator","Viva Pulse 관리자","공동 작업","S","W","N",False,"",
 "Viva Pulse 앱 설정 관리"),
("Insights Administrator","인사이트 관리자","공동 작업","S","W","P",False,"CONTENT",
 "M365 Insights 앱 관리자 권한. 임직원 협업·활동 데이터 접근"),
("Insights Analyst","Insights 분석가","공동 작업","S","R","P",False,"CONTENT",
 "Viva Insights 분석 기능 사용. 임직원 활동 패턴 데이터 조회"),
("Insights Business Leader","Insights 비즈니스 리더","읽기 전용","S","R","N",False,"",
 "Viva Insights 대시보드·인사이트 조회 및 공유"),

# ── 디바이스 ────────────────────────────────────────────────────────────────
("Intune Administrator","Intune 관리자","디바이스","S","W","P",True,"ESC-DEVICE",
 "Intune 전반 관리. 관리 디바이스에 스크립트·구성을 배포할 수 있어 엔드포인트 장악 경로"),
("Cloud Device Administrator","클라우드 디바이스 관리자","디바이스","T","W","P",True,"ESC-DEVICE",
 "Entra ID 디바이스 관리 및 BitLocker 키 조회 가능. 디바이스 접근 경로"),
("Microsoft Entra Joined Device Local Administrator","Microsoft Entra 조인 디바이스 로컬 관리자","디바이스","T","W","P",False,"ESC-DEVICE",
 "모든 Entra 조인 디바이스의 로컬 관리자 그룹에 추가됨. 엔드포인트 전체에 로컬 관리자 권한"),
("Windows 365 Administrator","Windows 365 관리자","디바이스","S","W","P",False,"ESC-DEVICE",
 "클라우드 PC 프로비저닝 및 전반 관리. 사용자 데스크톱 환경 접근 경로"),
("Windows Update Deployment Administrator","Windows 업데이트 배포 관리자","디바이스","S","W","N",False,"",
 "Windows Update 배포 전반 생성·관리"),
("Desktop Analytics Administrator","Desktop Analytics 관리자","디바이스","S","W","N",False,"",
 "데스크톱 관리 도구 및 서비스 접근·관리"),
("Printer Administrator","프린터 관리자","디바이스","S","W","N",False,"",
 "프린터 및 프린터 커넥터 전반 관리"),
("Printer Technician","프린터 기술자","디바이스","S","W","N",False,"",
 "프린터 등록·해제 및 상태 업데이트"),
("IoT Device Administrator","IoT 디바이스 관리자","디바이스","S","W","N",False,"",
 "IoT 디바이스 프로비저닝·수명주기·인증서 구성"),
("Network Administrator","네트워크 관리자","디바이스","S","W","N",False,"",
 "네트워크 위치 관리 및 엔터프라이즈 네트워크 설계 인사이트 검토"),
("Edge Administrator","Edge 관리자","기타","S","W","N",False,"",
 "Microsoft Edge의 Internet Explorer 모드 사이트 목록 관리"),
("Microsoft Hardware Warranty Administrator","Microsoft 하드웨어 보증 관리자","디바이스","S","W","N",False,"",
 "Microsoft 하드웨어 보증 청구·자격 전반 관리"),
("Microsoft Hardware Warranty Specialist","Microsoft 하드웨어 보증 전문가","디바이스","S","R","N",False,"",
 "Microsoft 하드웨어 보증 청구 생성·조회"),

# ── 읽기 전용 / 보고 ────────────────────────────────────────────────────────
("Global Reader","전역 읽기 권한자","읽기 전용","T","R","N",True,"",
 "전역 관리자가 볼 수 있는 모든 관리 기능·설정을 조회. 변경은 불가하나 M365·Entra 전반의 광범위한 정보가 노출됨"),
("Directory Reader","디렉터리 읽기 권한자","읽기 전용","T","R","N",False,"",
 "기본 디렉터리 정보 조회"),
("Reports Reader","보고서 읽기 권한자","읽기 전용","S","R","N",False,"",
 "사용 현황·활동·로그인 보고서 조회. 개인별 활동 데이터가 포함될 수 있음"),
("Usage Summary Reports Reader","사용 요약 보고서 읽기 권한자","읽기 전용","S","R","N",False,"",
 "사용 요약 보고서 및 채택 점수 조회 (개인 식별 정보 제외)"),
("Message Center Reader","메시지 센터 읽기 권한자","읽기 전용","S","R","N",False,"",
 "메시지 센터 알림 조회 및 Entra 사용자·그룹 읽기 전용 접근"),
("Message Center Privacy Reader","메시지 센터 개인 정보 읽기 권한자","읽기 전용","S","R","N",False,"",
 "메시지 센터의 개인정보·보안 메시지 조회. 전역 관리자와 함께 데이터 개인정보 메시지 열람 권한 보유"),
("Service Support Administrator","서비스 지원 관리자","읽기 전용","S","W","N",False,"",
 "서비스 요청 관리, 메시지 센터 조회, 서비스 상태 모니터링"),
("User Experience Success Manager","사용자 경험 성공 관리자","읽기 전용","S","R","N",False,"",
 "환경 인사이트·채택 점수·메시지 센터 조회"),
("Attribute Assignment Reader","특성 할당 읽기 권한자","읽기 전용","T","R","N",False,"",
 "개체의 사용자 지정 보안 특성 값 조회"),
("Attribute Definition Reader","특성 정의 읽기 권한자","읽기 전용","T","R","N",False,"",
 "사용자 지정 보안 특성 정의 조회"),
("Attribute Provisioning Reader","특성 프로비저닝 판독기","읽기 전용","T","R","N",True,"",
 "앱의 사용자 지정 보안 특성 프로비저닝 구성 조회"),
("Tenant Governance Reader","테넌트 거버넌스 판독기","읽기 전용","T","R","N",False,"",
 "테넌트 거버넌스 데이터 전체 조회"),
("Tenant Governance Relationship Manager","테넌트 거버넌스 관계 관리자","ID","T","W","N",False,"",
 "거버넌스 관계 시작·종료"),
("Tenant Governance Relationship Reader","테넌트 거버넌스 관계 판독기","읽기 전용","T","R","N",False,"",
 "테넌트 거버넌스 관계 및 관련 개체 조회"),
("Customer Delegated Admin Relationship Manager","고객 위임 관리자 관계 관리자","ID","T","W","P",False,"ESC-ROLE",
 "고객 테넌트의 GDAP 관계 전반 관리. 파트너 위임 관리 경로"),
("Agent Registry Administrator","에이전트 레지스트리 관리자","기타","S","W","N",False,"",
 "Entra ID 에이전트 레지스트리 서비스 전반 관리"),

# ── 청구 / 기타 ─────────────────────────────────────────────────────────────
("Billing Administrator","청구 관리자","청구","S","W","N",False,"",
 "구매·구독·서비스 요청 관리. 비용 발생 및 신규 서비스 활성화 가능"),
("License Administrator","라이선스 관리자","청구","S","W","N",False,"",
 "사용자·그룹 라이선스 할당·제거 및 사용 위치 편집"),
("Azure DevOps Administrator","Azure DevOps 관리자","기타","S","W","N",False,"",
 "Azure DevOps 정책·설정 관리"),
("Dragon Administrator","Dragon 관리자","기타","S","W","N",False,"",
 "Microsoft Dragon 관리 센터 전반 관리"),
]


def grade_of(scope, op, sens):
    """8조합 → A1/A2/A3 매핑"""
    if sens == "P":
        return "A3"
    if op == "R":
        return "A1" if scope == "S" else "A2"
    return "A2" if scope == "S" else "A3"


COMBO_MAP = [
    ("SRN", "S", "R", "N", "A1", "서비스 한정 · 읽기 전용 · 민감 접근 없음"),
    ("TRN", "T", "R", "N", "A2", "테넌트 전역 · 읽기 전용 · 민감 접근 없음"),
    ("SWN", "S", "W", "N", "A2", "서비스 한정 · 구성 변경 · 민감 접근 없음"),
    ("TWN", "T", "W", "N", "A3", "테넌트 전역 · 구성 변경 · 민감 접근 없음"),
    ("SRP", "S", "R", "P", "A3", "서비스 한정 · 읽기 전용 · 민감 접근 있음"),
    ("TRP", "T", "R", "P", "A3", "테넌트 전역 · 읽기 전용 · 민감 접근 있음"),
    ("SWP", "S", "W", "P", "A3", "서비스 한정 · 구성 변경 · 민감 접근 있음"),
    ("TWP", "T", "W", "P", "A3", "테넌트 전역 · 구성 변경 · 민감 접근 있음"),
]

AXIS_LABEL = {
    "S": "서비스·기능 한정", "T": "테넌트 전역",
    "R": "읽기 전용", "W": "구성 변경",
    "N": "없음", "P": "있음",
}

DOC_BASE = "https://learn.microsoft.com/ko-kr/entra/identity/role-based-access-control/permissions-reference"


def anchor(en_name):
    return DOC_BASE + "#" + en_name.lower().replace(" ", "-").replace("(", "").replace(")", "").replace(".", "")


def build():
    """판정을 적용한 역할 레코드 리스트 반환"""
    out = []
    seen = set()
    for en, ko, cat, scope, op, sens, priv, reason, risk in ROLES:
        if en in seen:
            raise ValueError("중복 역할: " + en)
        seen.add(en)
        g = grade_of(scope, op, sens)
        out.append(dict(
            en=en, ko=ko, category=cat,
            scope=scope, op=op, sens=sens,
            combo=scope + op + sens,
            grade=g,
            risk=GRADE_DEF[g]["risk"],
            privileged=priv,
            sensReason=SENS_REASON.get(reason, "") if reason else "",
            sensReasonCode=reason,
            keyRisk=risk,
            doc=anchor(en),
        ))
    out.sort(key=lambda r: (["A3", "A2", "A1"].index(r["grade"]), r["ko"]))
    return out


if __name__ == "__main__":
    import io, sys, collections
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    rows = build()
    print("총 역할:", len(rows))
    print("등급 분포:", dict(collections.Counter(r["grade"] for r in rows)))
    print("조합 분포:", dict(collections.Counter(r["combo"] for r in rows)))
    print("MS 권한있는역할:", sum(1 for r in rows if r["privileged"]))
    print("범주 분포:", dict(collections.Counter(r["category"] for r in rows)))
    # 사용자 예시 검증
    for k in ["Power Platform Administrator", "Global Reader", "AI Administrator",
              "Compliance Administrator", "Security Administrator"]:
        r = next(x for x in rows if x["en"] == k)
        print(f'  {k:34s} {r["combo"]} → {r["grade"]} {r["risk"]}')
