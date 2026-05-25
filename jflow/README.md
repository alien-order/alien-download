# jflow — Spring + MyBatis + JSP 프로세스 흐름 추출기

AI 없이 **순수 정적 분석**으로, Java 웹앱의 한 기능을 골라 호출하면
**JSP → @Controller → @Service → @Mapper → SQL → 테이블** 흐름을 펼쳐서 보여주는 단일 파일 도구.

```
[URL] POST /order/save
`-- [CTRL] OrderController.save
    |-- [VIEW] view: redirect:/order/list
    `-- [SVC] OrderServiceImpl.save        (OrderService 인터페이스의 구현)
        `-- [MAP] OrderMapper.insertOrder
            `-- [SQL] INSERT -> T_ORDER
```

## 설치 (옮길 PC에서)

```bash
pip install javalang        # 유일한 의존성 (순수 파이썬)
```

`jflow.py` 파일 하나만 있으면 됩니다. 회사 소스가 있는 PC로 `jflow.py` + `pip install javalang` 만 옮기면 끝.

## 사용법

소스 루트(`--src`)는 여러 번 지정 가능 (java / mapper xml / jsp 가 다른 폴더면 각각 추가).

```bash
# 1) 인덱스 요약 — 컨트롤러/매퍼/URL 몇 개 잡혔는지 먼저 확인
python jflow.py stats --src C:\proj\src --src C:\proj\webapp --errors

# 2) URL → 컨트롤러 매핑 전체 목록 (어디서 시작할지 고르기)
python jflow.py urls  --src C:\proj --grep order

# 3) 흐름 추적 — 진입점 3가지 중 하나
python jflow.py flow --src C:\proj --url    /order/save
python jflow.py flow --src C:\proj --method com.demo.web.OrderController.save
python jflow.py flow --src C:\proj --jsp    order/list.jsp     # 경로 뒤쪽만 맞아도 됨

# 4) Mermaid 다이어그램으로 (붙여넣으면 그림으로 렌더)
python jflow.py flow --src C:\proj --url /order/save --mermaid

# 5) JSP의 ${ctx} 같은 컨텍스트 경로가 URL 앞에 붙는 경우
python jflow.py flow --src C:\proj --jsp list.jsp --context-path /myapp
```

> Windows PowerShell/cmd 에서는 `--url /order/save` 그대로 됩니다.
> (Git Bash 는 `/order/...` 를 윈도우 경로로 바꿔버리니 PowerShell 권장.)

## 무엇을 어떻게 해석하나 (전부 결정론적)

| 단계 | 근거 |
|------|------|
| URL → 메서드 | 클래스 `@RequestMapping` + 메서드 `@RequestMapping/@GetMapping/@PostMapping/...` 값 결합 |
| 필드 호출 | `someService.foo()` → 선언된 필드 타입 → 빈 타입 |
| 인터페이스 → 구현 | 인터페이스 타입 → 그걸 `implements` 한 `@Service/@Repository/@Component` 클래스 |
| 매퍼 → SQL | `@Mapper` 인터페이스 메서드 == MyBatis XML `<select|insert|update|delete>` id (XML `namespace` == 인터페이스 FQN). `@Select/@Insert` 어노테이션 SQL도 인식 |
| SQL → 테이블 | `FROM/JOIN/INTO/UPDATE` 파싱, `<include refid>` 펼침, `#{}`/`${}` 치환 |
| 뷰 | 컨트롤러 메서드의 `return "view/name"` / `new ModelAndView("...")` |

## 한계 (정직하게)

- ✅ 잘 됨: 설정/어노테이션 기반 매핑, DI 스파인(Controller→Service→Mapper), MyBatis SQL→테이블
- ⚠️ 부분적: 메서드 체인 `a.getB().c()`, 동적 SQL의 일부, 같은 클래스 private 헬퍼는 따라가되 외부 라이브러리 호출은 생략
- ❌ 거의 안 됨: 리플렉션, 런타임 문자열로 만든 동적 디스패치, JSP 스크립틀릿 내부 분기 로직
- URL 매칭이 잘 안 되면 `urls` 로 목록을 본 뒤 `--method` 로 컨트롤러 메서드부터 직접 시작하세요.
- 파싱 실패한 `.java` 는 건너뛰고 `stats --errors` 로 목록을 보여줍니다(전체 중단 안 함).

## 자체 검증

`_selftest/` 에 합성 미니 앱(Spring+MyBatis+JSP)이 들어 있습니다. 동작 확인:

```bash
python jflow.py flow --src _selftest --url /order/save
```
