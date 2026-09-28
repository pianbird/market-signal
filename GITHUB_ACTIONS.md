# GitHub Actions 실행

다음 파일을 GitHub 저장소의 기본 브랜치에 함께 커밋하고 푸시하세요.

- `market_integrated_analysis_v2.5.py`
- `requirements.txt`
- `.github/workflows/market-analysis.yml`

GitHub의 **Actions → Market analysis → Run workflow**에서 수동 실행합니다.
실행 상세 화면의 **Artifacts**에서 엑셀 보고서를 다운로드합니다. 보관 기간은 30일입니다.
워크플로는 Ubuntu / Python 3.11을 사용하며 시간대는 Asia/Seoul입니다.
자동 예약 실행은 설정하지 않았습니다.

## 선택적 외부 연동

저장소 **Settings → Secrets and variables → Actions → New repository secret**에
사용할 연동의 값을 모두 등록하세요. 등록한 연동은 실행 시 실제 전송/등록됩니다.

| 연동 | 필요한 Secrets |
| --- | --- |
| Telegram | `MY_BOT_TOKEN`, `MY_CHAT_ID` |
| Notion | `NOTION_API_KEY`, `NOTION_DATABASE_ID` |

Notion 연동에는 대상 데이터베이스에 대한 통합 접근 권한도 필요합니다.
Secrets를 등록하지 않아도 엑셀 분석은 실행되며 외부 연동은 건너뜁니다.
토큰이나 `.env` 파일은 저장소에 커밋하지 마세요.

## 저장 경로 및 실패 처리

- `OUTPUT_DIR` 환경변수가 있으면 해당 폴더를 생성해 사용합니다.
- 없으면 `/content` 폴더가 존재하는 환경에서는 `/content`, 그 외에는 현재 작업 폴더에 저장합니다.
- Actions에서는 `output/`에 저장하고 아티팩트로 업로드합니다.
- 분석 데이터가 전혀 없거나 엑셀 저장에 실패하면 실행이 실패합니다.
- 외부 시세 서비스의 네트워크 제한이나 장애로 수집이 실패할 수 있습니다.
- 기존 동작대로 일부 ETF만 수집된 경우 수집된 종목으로 보고서를 생성합니다.
- 선택적 외부 연동 실패는 엑셀 분석 작업의 실패로 처리하지 않습니다. 실행 로그에서 확인하세요.

로컬 실행:

```sh
python -m pip install -r requirements.txt
python market_integrated_analysis_v2.5.py
```
