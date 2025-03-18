# PDF-X: PDF 문서 분석 및 데이터 추출 도구

이 도구는 PDF 파일에서 텍스트를 추출하고, 네이버 클로바 OCR API와 OpenAI API를 활용하여 데이터를 분석한 후 Excel 파일로 변환하는 기능을 제공합니다.

## 주요 기능

1. **폴더 기반 자동 처리**: `source` 폴더 내 하위 폴더에 있는 PDF 파일을 자동으로 감지하여 처리합니다.
2. **PDF 변환 및 처리**: PDF 파일을 이미지로 변환하여 각 페이지를 개별적으로 처리합니다.
3. **OCR 텍스트 추출**: 네이버 클로바 OCR API를 사용하여 이미지에서 텍스트를 추출합니다.
4. **데이터 분석**: OpenAI API(gpt-4o-mini 모델)를 사용하여 추출된 텍스트에서 주요 데이터를 분석합니다.
5. **데이터 구조화**: 분석된 데이터를 구조화하여 Excel 파일로 변환합니다.
6. **날짜별 결과 관리**: 처리된 결과는 날짜별로 관리되며, 하나의 Excel 파일에 누적됩니다.
7. **날짜별 로그 관리**: 로그는 날짜별로 별도 파일에 저장됩니다.
8. **병합 파일 관리**: 처리된 파일은 원본 폴더 구조를 유지하며 `merged` 폴더에 병합되어 저장됩니다.
9. **다중 처리 최적화**: ThreadPoolExecutor를 사용한 병렬 처리로 여러 파일을 효율적으로 처리합니다.
10. **재시도 메커니즘**: 네트워크 요청에 재시도 로직을 적용하여 일시적인 오류를 방지합니다.

## 시스템 요구사항

- **Python 버전**: 3.13.2 이상
- **운영체제**: Windows, macOS, Linux

## 설치 방법

1. 필요한 패키지 설치:
   ```
   pip install -r requirements.txt
   ```

2. 환경 설정:
   - 프로젝트 루트 디렉토리에 `.env` 파일을 생성하고 다음과 같이 설정합니다:
     ```
     # 네이버 클로바 OCR API 설정
     CLOVA_OCR_SECRET_KEY=your_clova_ocr_secret_key_here
     CLOVA_OCR_APIGW_INVOKE_URL=https://paper.ncloud.com/api/v2/document/ocr
     
     # OpenAI API 설정
     OPENAI_API_KEY=your_openai_api_key_here
     
     # 기타 설정 (선택 사항)
     MAX_FILE_SIZE=16777216  # 최대 파일 크기 (바이트 단위, 기본값: 16MB)
     SOURCE_DIR=source  # 소스 폴더 경로
     MERGED_DIR=merged  # 병합 폴더 경로
     RESULT_DIR=result  # 결과 폴더 경로
     TEMP_DIR=temp  # 임시 폴더 경로
     LOG_DIR=logs  # 로그 폴더 경로
     LOG_LEVEL=INFO  # 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
     ```

## 사용 방법

1. `source` 폴더 내에 하위 폴더를 생성하고 PDF 파일을 업로드합니다.
   - 예: `source/테스트1/문서1.pdf`, `source/테스트2/문서2.pdf`
   
2. 다음 명령어로 프로그램을 실행합니다:
   ```
   python main.py
   ```
   
3. 프로그램이 자동으로 파일을 감지하고 처리합니다.
   - PDF 파일은 이미지로 변환됩니다.
   - 이미지는 OCR 처리되어 텍스트가 추출됩니다.
   - 추출된 텍스트는 OpenAI API를 통해 분석됩니다.
   - 분석 결과는 Excel 파일로 저장됩니다.

4. 처리 결과는 `result/output_YYYYMMDD.xlsx` 파일로 저장됩니다.
5. 로그는 `logs/log_YYYYMMDD.log` 파일에 저장됩니다.
6. 병합된 JSONL 파일은 `merged/[하위폴더명]/[하위폴더명].jsonl` 형태로 저장됩니다.

## 폴더 구조

- `source/`: 처리할 PDF 파일을 포함하는 하위 폴더를 생성하는 곳입니다.
- `merged/`: 원본 폴더 구조를 유지하며 병합된 JSONL 파일이 저장되는 폴더입니다.
- `result/`: 처리 결과가 저장되는 폴더입니다.
- `temp/YYYYMMDD/`: 날짜별로 임시 파일이 저장되는 폴더입니다 (자동 생성).
- `logs/`: 로그 파일이 저장되는 폴더입니다.

## 지원 파일 형식

- PDF: `.pdf`

## 로그

프로그램 실행 중 발생하는 모든 이벤트는 `logs/log_YYYYMMDD.log` 파일에 기록됩니다. 로그는 날짜별로 별도 파일에 저장됩니다.

## 결과 파일

처리 결과는 `result/output_YYYYMMDD.xlsx` 파일에 저장됩니다. 같은 날짜에 처리된 모든 파일의 결과가 하나의 Excel 파일에 누적됩니다.

Excel 파일에는 다음 정보가 포함됩니다:
- 접수일자: 문서에서 추출된 접수일자 (YYYY-MM-DD 형식)
- 세목: 민간경상보조, 민간위탁사업비, 사업출연금 중 하나
- 세부사업명: 문서에서 추출된 세부사업명
- 내역사업명: 문서에서 추출된 사업명 정보
- 기관명: 내역사업명 옆에 표시되는 예금주 정보
- 신뢰도: OCR 결과의 평균 신뢰도 (%)
- 금액 정보: 연간집행계획액, 기수령액, 월집행계획액, 전월이월액, 당월신청액, 누계 등
- 은행 정보: 은행명, 계좌번호, 예금주

## 환경 변수 설정

`.env` 파일에서 다음과 같은 환경 변수를 설정할 수 있습니다:

| 환경 변수 | 설명 | 기본값 |
|------------|-------------|---------|
| CLOVA_OCR_SECRET_KEY | 네이버 클로바 OCR API 비밀 키 | (필수) |
| CLOVA_OCR_APIGW_INVOKE_URL | 네이버 클로바 OCR API Gateway Invoke URL | https://paper.ncloud.com/api/v2/document/ocr |
| OPENAI_API_KEY | OpenAI API 키 | (필수) |
| MAX_FILE_SIZE | 최대 파일 크기 (바이트 단위) | 16777216 (16MB) |
| SOURCE_DIR | 소스 폴더 경로 | source |
| MERGED_DIR | 병합 폴더 경로 | merged |
| RESULT_DIR | 결과 폴더 경로 | result |
| TEMP_DIR | 임시 폴더 경로 | temp |
| LOG_DIR | 로그 폴더 경로 | logs |
| LOG_LEVEL | 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL) | INFO |

## 의존성 패키지

```
# 기본 패키지
watchdog==6.0.0  # Python 3.13과 호환되는 최신 버전
requests==2.31.0  # HTTP 요청 처리
openpyxl==3.1.2  # Excel 파일 처리
tqdm==4.66.1  # 진행률 표시
python-dotenv==1.0.0  # 환경 변수 관리
PyMuPDF==1.25.3  # PDF 처리 (Python 3.13과 호환)
Pillow==11.1.0  # 이미지 처리 (Python 3.13과 호환)
openai==1.12.0  # OpenAI API 사용 (gpt-4o-mini 모델 지원)

# 선택적 패키지
Flask==3.0.0  # 웹 프레임워크 (선택적) 
```

## API 정보

### 네이버 클로바 OCR API
이 도구는 네이버 클라우드 플랫폼의 클로바 OCR API를 사용합니다. API에 대한 자세한 정보는 다음 링크에서 확인할 수 있습니다:
- [네이버 클라우드 플랫폼 - Clova OCR](https://www.ncloud.com/product/aiService/ocr)
- [API 문서](https://api.ncloud-docs.com/docs/ai-application-service-ocr)

### OpenAI API
이 도구는 OpenAI의 GPT 모델(gpt-4o-mini)을 사용하여 OCR 결과를 분석합니다. API에 대한 자세한 정보는 다음 링크에서 확인할 수 있습니다:
- [OpenAI API](https://platform.openai.com/)
- [API 문서](https://platform.openai.com/docs/api-reference)
- [gpt-4o-mini 모델 정보](https://platform.openai.com/docs/models/gpt-4o-mini)

## 주요 기능 상세

### 1. 문서 처리 워크플로우
- **PDF 파일 감지**: 소스 폴더 내 PDF 파일을 자동으로 감지
- **이미지 변환**: PDF 각 페이지를 고품질 이미지로 변환
- **OCR 처리**: 네이버 클로바 OCR API를 통한 텍스트 추출
- **데이터 분석**: OpenAI API를 통한 추출 데이터 구조화
- **결과 저장**: 분석 결과를 Excel 파일로 저장

### 2. 성능 최적화
- **병렬 처리**: ThreadPoolExecutor를 사용한 다중 파일 동시 처리
- **재시도 메커니즘**: 네트워크 요청 실패 시 자동 재시도 (최대 3회)
- **진행률 표시**: tqdm을 활용한 실시간 작업 진행률 표시
- **로그 관리**: 로그 파일 자동 생성 및 날짜별 관리

### 3. 오류 처리
- **예외 처리**: 모든 주요 기능에 예외 처리 로직 적용
- **임시 파일 정리**: 작업 완료 후 임시 파일 자동 정리
- **오류 로깅**: 상세한 오류 정보를 로그 파일에 기록

## 주의사항

- 네이버 클로바 OCR API와 OpenAI API는 유료 서비스이므로, 사용량에 따라 비용이 발생할 수 있습니다.
- API 호출 시 네트워크 상태에 따라 지연이 발생할 수 있습니다.
- 이미지 품질이 낮은 경우 OCR 인식률이 저하될 수 있습니다.
- PDF 파일의 경우 페이지 수에 따라 처리 시간이 길어질 수 있습니다.

## 문제 해결

### Python 3.13 호환성 문제

Python 3.13에서는 스레딩 모듈의 변경으로 인해 일부 패키지에서 호환성 문제가 발생할 수 있습니다. 이 프로젝트는 다음과 같은 패키지 버전을 사용하여 Python 3.13과의 호환성 문제를 해결했습니다:

- **watchdog 6.0.0**: Python 3.13의 스레딩 모듈 변경에 대응하여 `'handle' must be a _ThreadHandle` 오류를 해결합니다.
- **Pillow 11.1.0**: Python 3.13과 호환되는 버전으로, 이전 버전에서 발생하는 `KeyError: '__version__'` 오류를 해결합니다.
- **PyMuPDF 1.25.3**: Python 3.13과 호환되는 버전으로, PDF 처리 기능을 제공합니다.

### 네트워크 관련 문제

- **API 호출 오류**: API 호출 시 오류가 발생한 경우, 자동으로 최대 3회까지 재시도합니다.
- **타임아웃 오류**: API 호출 타임아웃은 기본 60초로 설정되어 있으며, 필요에 따라 코드에서 조정할 수 있습니다.

### 기타 알려진 문제

- **대용량 PDF 처리**: 매우 큰 PDF 파일(100MB 이상)의 경우 메모리 사용량이 증가할 수 있으므로 주의가 필요합니다.
- **비표준 PDF**: 일부 비표준 PDF 파일은 제대로 처리되지 않을 수 있습니다. 가능한 표준 PDF 파일을 사용하세요.

## 업데이트 내역

### 2025-03-18
- 코드 최적화: 병렬 처리 로직 개선
- 버그 수정: OCR 결과 병합 및 분석 과정에서 발생하는 오류 수정
- 기능 추가: 진행률 표시 개선 및 작업 수 계산 로직 추가

### 2025-03-12
- 최초 버전 릴리스 