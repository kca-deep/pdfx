# 네이버 클로바 OCR API와 OpenAI API를 활용한 문서 분석 도구

이 도구는 이미지 및 PDF 파일에서 텍스트를 추출하고, OpenAI API를 활용하여 데이터를 분석한 후 Excel 파일로 변환하는 기능을 제공합니다.

## 주요 기능

1. **파일 감지 및 자동 처리**: `source` 폴더에 새 이미지 또는 PDF 파일이 업로드되면 자동으로 감지하여 처리합니다.
2. **PDF 변환 및 처리**: PDF 파일을 이미지로 변환하여 각 페이지를 개별적으로 처리합니다.
3. **OCR 텍스트 추출**: 네이버 클로바 OCR API를 사용하여 이미지에서 텍스트를 추출합니다.
4. **데이터 분석**: OpenAI API를 사용하여 추출된 텍스트에서 금액 데이터와 은행 정보를 분석합니다.
5. **데이터 구조화**: 분석된 데이터를 구조화하여 Excel 파일로 변환합니다.
6. **날짜별 결과 관리**: 처리된 결과는 날짜별 폴더에 저장되며, 하나의 Excel 파일에 누적됩니다.
7. **날짜별 로그 관리**: 로그는 날짜별로 별도 파일에 저장됩니다.

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
     LOG_LEVEL=INFO  # 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
     ```
   - 또는 환경 변수로 직접 설정할 수도 있습니다:
     - Windows: 
       ```
       set CLOVA_OCR_SECRET_KEY=your_api_key_here
       set CLOVA_OCR_APIGW_INVOKE_URL=your_api_gateway_url_here
       set OPENAI_API_KEY=your_openai_api_key_here
       ```
     - Linux/Mac: 
       ```
       export CLOVA_OCR_SECRET_KEY=your_api_key_here
       export CLOVA_OCR_APIGW_INVOKE_URL=your_api_gateway_url_here
       export OPENAI_API_KEY=your_openai_api_key_here
       ```

## 사용 방법

1. `source` 폴더에 이미지 또는 PDF 파일을 업로드합니다.
2. 다음 명령어로 프로그램을 실행합니다:
   ```
   python main.py
   ```
3. 프로그램이 자동으로 파일을 감지하고 처리합니다.
4. 처리 결과는 `result/YYYYMMDD` 폴더에 `output.xlsx` 파일로 저장됩니다.
5. 로그는 `logs/pdfx_YYYYMMDD.log` 파일에 저장됩니다.

### 기존 파일 처리

프로그램 실행 시 `source` 폴더에 이미 존재하는 파일도 자동으로 처리됩니다.

## 폴더 구조

- `source/`: 처리할 이미지 또는 PDF 파일을 이 폴더에 업로드합니다.
- `result/YYYYMMDD/`: 날짜별로 처리 결과가 저장되는 폴더입니다.
- `temp/YYYYMMDD/`: 날짜별로 임시 파일이 저장되는 폴더입니다 (자동 생성).
- `logs/`: 로그 파일이 저장되는 폴더입니다.

## 지원 파일 형식

- 이미지: `.jpg`, `.jpeg`, `.png`, `.tiff`, `.tif`, `.bmp`
- PDF: `.pdf`

## 로그

프로그램 실행 중 발생하는 모든 이벤트는 `logs/pdfx_YYYYMMDD.log` 파일에 기록됩니다. 로그는 날짜별로 별도 파일에 저장됩니다.

## 결과 파일

처리 결과는 `result/YYYYMMDD/output.xlsx` 파일에 저장됩니다. 같은 날짜에 처리된 모든 파일의 결과가 하나의 Excel 파일에 누적됩니다.

Excel 파일에는 다음 정보가 포함됩니다:
- 원본 파일명
- 처리 시간
- 페이지 번호 (PDF의 경우)
- 금액 데이터 (연간집행계획액, 기수령액, 월집행계획액, 전월이월액, 당월신청액, 누계 등)
- 은행 정보 (은행명, 계좌번호, 예금주)

## 환경 변수 설정

`.env` 파일에서 다음과 같은 환경 변수를 설정할 수 있습니다:

| 환경 변수 | 설명 | 기본값 |
|------------|-------------|---------|
| CLOVA_OCR_SECRET_KEY | 네이버 클로바 OCR API 비밀 키 | (필수) |
| CLOVA_OCR_APIGW_INVOKE_URL | 네이버 클로바 OCR API Gateway Invoke URL | https://paper.ncloud.com/api/v2/document/ocr |
| OPENAI_API_KEY | OpenAI API 키 | (필수) |
| MAX_FILE_SIZE | 최대 파일 크기 (바이트 단위) | 16777216 (16MB) |
| SOURCE_DIR | 소스 폴더 경로 | source |
| RESULT_DIR | 결과 폴더 경로 | result |
| TEMP_DIR | 임시 폴더 경로 | temp |
| LOG_LEVEL | 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL) | INFO |

## API 정보

### 네이버 클로바 OCR API
이 도구는 네이버 클라우드 플랫폼의 클로바 OCR API를 사용합니다. API에 대한 자세한 정보는 다음 링크에서 확인할 수 있습니다:
- [네이버 클라우드 플랫폼 - Clova OCR](https://www.ncloud.com/product/aiService/ocr)
- [API 문서](https://api.ncloud-docs.com/docs/ai-application-service-ocr)

### OpenAI API
이 도구는 OpenAI의 GPT 모델을 사용하여 OCR 결과를 분석합니다. API에 대한 자세한 정보는 다음 링크에서 확인할 수 있습니다:
- [OpenAI API](https://platform.openai.com/)
- [API 문서](https://platform.openai.com/docs/api-reference)

## 주의사항

- 네이버 클로바 OCR API와 OpenAI API는 유료 서비스이므로, 사용량에 따라 비용이 발생할 수 있습니다.
- API 호출 시 네트워크 상태에 따라 지연이 발생할 수 있습니다.
- 이미지 품질이 낮은 경우 OCR 인식률이 저하될 수 있습니다.
- PDF 파일의 경우 페이지 수에 따라 처리 시간이 길어질 수 있습니다. 