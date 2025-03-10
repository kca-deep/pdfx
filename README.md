# 네이버 클로바 OCR API를 활용한 표 인식 및 텍스트 추출 도구

이 도구는 이미지 파일에서 표를 인식하고 텍스트를 추출하여 Excel 파일로 변환하는 기능을 제공합니다.

> **참고**: 현재 버전에서는 PDF 파일 지원이 비활성화되어 있습니다. 이미지 파일(.jpg, .jpeg, .png, .tiff, .tif, .bmp)만 처리 가능합니다.

## 주요 기능

1. **파일 감지 및 자동 처리**: `source` 폴더에 새 이미지 파일이 업로드되면 자동으로 감지하여 처리합니다.
2. **표 인식 및 텍스트 추출**: 네이버 클로바 OCR API를 사용하여 이미지 내의 표를 인식하고 텍스트를 추출합니다.
3. **데이터 구조화**: 추출된 표 데이터를 구조화하여 Excel 파일로 변환합니다.
4. **결과 저장**: 처리된 결과는 `result` 폴더에 타임스탬프가 포함된 Excel 파일로 저장됩니다.

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
     
     # 기타 설정 (선택 사항)
     MAX_FILE_SIZE=16777216  # 최대 파일 크기 (바이트 단위, 기본값: 16MB)
     LOG_LEVEL=INFO  # 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL)
     ```
   - 또는 환경 변수로 직접 설정할 수도 있습니다:
     - Windows: 
       ```
       set CLOVA_OCR_SECRET_KEY=your_api_key_here
       set CLOVA_OCR_APIGW_INVOKE_URL=your_api_gateway_url_here
       ```
     - Linux/Mac: 
       ```
       export CLOVA_OCR_SECRET_KEY=your_api_key_here
       export CLOVA_OCR_APIGW_INVOKE_URL=your_api_gateway_url_here
       ```

## 사용 방법

1. `source` 폴더에 이미지 파일을 업로드합니다.
2. 다음 명령어로 프로그램을 실행합니다:
   ```
   python main.py
   ```
3. 프로그램이 자동으로 파일을 감지하고 처리합니다.
4. 처리 결과는 `result` 폴더에 저장됩니다.

### API 테스트 (test.py)

1. 테스트 이미지 파일을 준비합니다 (`test_image.png`).
2. 다음 명령어로 API 테스트를 실행합니다:
   ```
   python test.py
   ```
3. 테스트 결과는 콘솔에 출력되며, 상세 결과는 `test_result.json` 파일에 저장됩니다.

## 폴더 구조

- `source/`: 처리할 이미지 파일을 이 폴더에 업로드합니다.
- `result/`: 처리 결과가 저장되는 폴더입니다.
- `temp/`: 임시 파일이 저장되는 폴더입니다 (자동 생성).

## 지원 파일 형식

- 이미지: `.jpg`, `.jpeg`, `.png`, `.tiff`, `.tif`, `.bmp`

## 로그

프로그램 실행 중 발생하는 모든 이벤트는 `script.log` 파일에 기록됩니다.

## 환경 변수 설정

`.env` 파일에서 다음과 같은 환경 변수를 설정할 수 있습니다:

| 환경 변수 | 설명 | 기본값 |
|------------|-------------|---------|
| CLOVA_OCR_SECRET_KEY | 네이버 클로바 OCR API 비밀 키 | (필수) |
| CLOVA_OCR_APIGW_INVOKE_URL | 네이버 클로바 OCR API Gateway Invoke URL | https://paper.ncloud.com/api/v2/document/ocr |
| OPENAI_API_KEY | OpenAI API 키 (필요한 경우) | (선택 사항) |
| MAX_FILE_SIZE | 최대 파일 크기 (바이트 단위) | 16777216 (16MB) |
| SOURCE_DIR | 소스 폴더 경로 | source |
| RESULT_DIR | 결과 폴더 경로 | result |
| TEMP_DIR | 임시 폴더 경로 | temp |
| LOG_LEVEL | 로그 레벨 (DEBUG, INFO, WARNING, ERROR, CRITICAL) | INFO |

## 네이버 클로바 OCR API 정보

이 도구는 네이버 클라우드 플랫폼의 클로바 OCR API를 사용합니다. API에 대한 자세한 정보는 다음 링크에서 확인할 수 있습니다:
- [네이버 클라우드 플랫폼 - Clova OCR](https://www.ncloud.com/product/aiService/ocr)
- [API 문서](https://api.ncloud-docs.com/docs/ai-application-service-ocr)

## 주의사항

- 네이버 클로바 OCR API는 유료 서비스이므로, 사용량에 따라 비용이 발생할 수 있습니다.
- API 호출 시 네트워크 상태에 따라 지연이 발생할 수 있습니다.
- 이미지 품질이 낮거나 복잡한 표 구조의 경우 인식률이 저하될 수 있습니다. 