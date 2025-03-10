import os
import sys
import json
import time
import base64
import logging
import requests
import mimetypes
from pathlib import Path
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv
import openpyxl
import fitz  # PyMuPDF
from PIL import Image
import io
import traceback

# import openai  # OpenAI API 사용을 위한 라이브러리 추가 - requests 라이브러리로 대체

# 환경 변수 로드
load_dotenv(override=True)  # override=True로 설정하여 기존 환경 변수를 덮어씁니다.

# 로그 디렉토리 확인 및 생성
log_dir = Path("logs")
if not log_dir.exists():
    log_dir.mkdir(parents=True, exist_ok=True)

# 날짜 기반 로그 파일명 생성 (하루 단위)
current_date = datetime.now().strftime("%Y%m%d")
log_filename = log_dir / f"pdfx_{current_date}.log"

# 로깅 설정
logging.basicConfig(
    level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(str(log_filename), mode="a", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

logging.info(f"로그 파일 경로: {log_filename}")

# 네이버 클로바 OCR API 설정
CLOVA_OCR_API_URL = os.getenv("CLOVA_OCR_APIGW_INVOKE_URL")
CLOVA_OCR_API_SECRET = os.getenv("CLOVA_OCR_SECRET_KEY")

# OpenAI API 설정
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"
# openai.api_key = OPENAI_API_KEY  # requests 라이브러리로 대체

# 설정 로그
logging.info(f"CLOVA OCR API URL: {CLOVA_OCR_API_URL}")
if CLOVA_OCR_API_SECRET:
    logging.info(
        f"CLOVA OCR API SECRET: {CLOVA_OCR_API_SECRET[:4]}...{CLOVA_OCR_API_SECRET[-4:]} (길이: {len(CLOVA_OCR_API_SECRET)})"
    )

# 기타 설정
MAX_FILE_SIZE = int(os.getenv("MAX_FILE_SIZE", 16777216))  # 기본값: 16MB
TEMP_DIR = os.getenv("TEMP_DIR", "temp")
SOURCE_DIR = os.getenv("SOURCE_DIR", "source")
RESULT_DIR = os.getenv("RESULT_DIR", "result")

# 지원하는 파일 확장자
SUPPORTED_IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"]
SUPPORTED_PDF_EXTENSION = ".pdf"
ALL_SUPPORTED_EXTENSIONS = SUPPORTED_IMAGE_EXTENSIONS + [SUPPORTED_PDF_EXTENSION]

# 필요한 디렉토리 생성
for directory in [TEMP_DIR, SOURCE_DIR, RESULT_DIR]:
    os.makedirs(directory, exist_ok=True)


def analyze_ocr_with_openai(ocr_result):
    """OpenAI API를 사용하여 OCR 결과 분석 (requests 라이브러리 사용)"""
    if not OPENAI_API_KEY:
        logging.error("OpenAI API 키가 설정되지 않았습니다.")
        return None

    try:
        logging.info("OpenAI API를 사용하여 OCR 결과 분석 시작")

        # OCR 결과에서 텍스트 추출
        all_text = ""
        if "images" in ocr_result and ocr_result["images"]:
            image_result = ocr_result["images"][0]

            # 텍스트 필드 추출
            if "fields" in image_result:
                for field in image_result["fields"]:
                    if "inferText" in field:
                        # 텍스트 좌표 정보 추출
                        vertices = field.get("boundingPoly", {}).get("vertices", [])
                        if vertices:
                            y_coord = vertices[0].get("y", 0)
                            text = field.get("inferText", "")
                            all_text += f"행 {int(y_coord/10)}: {text}\n"

        # 분석할 텍스트가 없는 경우
        if not all_text:
            logging.warning("분석할 텍스트가 없습니다.")
            return None

        # OpenAI API 요청 프롬프트 작성
        prompt = f"""
다음은 OCR로 추출한 텍스트입니다. 이 텍스트에서 다음 정보를 추출해주세요:

1. 5~7행에 있는 금액 데이터: 연간집행, 기수령액, 계획액, 월집행, 이월액, 전월, 신청액, 당월, 누계
2. 22~24행에 있는 은행 정보: 은행명, 계좌번호, 예금주

추출한 정보를 다음 JSON 형식으로 반환해주세요:
{{
    "금액_데이터": [
        {{
            "항목": "연간집행계획액",
            "금액": "금액"
        }},
        {{
            "항목": "기수령액",
            "금액": "금액"
        }},
        {{
            "항목": "월집행계획액",
            "금액": "금액"
        }},
        {{
            "항목": "전월이월액",
            "금액": "금액"
        }},
        {{
            "항목": "당월신청액",
            "금액": "금액"
        }},
        {{
            "항목": "누계",
            "금액": "금액"
        }}
    ],
    "은행_정보": [
        {{
            "항목": "은행명",
            "값": "은행명"
        }},
        {{
            "항목": "계좌번호",
            "값": "계좌번호"
        }},
        {{
            "항목": "예금주",
            "값": "예금주"
        }}
    ]
}}
OCR 텍스트:
{all_text}
"""

        # OpenAI API 요청 데이터 준비
        payload = {
            "model": "gpt-4o-mini",  # 모델 지정
            "messages": [
                {
                    "role": "system",
                    "content": "당신은 OCR 결과에서 정보를 추출하는 전문가입니다.",
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.3,  # 낮은 temperature로 일관된 결과 유도
            "max_tokens": 1000,
        }

        # API 요청 헤더
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {OPENAI_API_KEY}",
        }

        # OpenAI API 호출 (requests 사용)
        logging.info("OpenAI API 요청 전송 중...")
        response = requests.post(OPENAI_API_URL, headers=headers, json=payload)

        # 응답 확인
        if response.status_code == 200:
            response_data = response.json()

            # 응답 추출
            if "choices" in response_data and len(response_data["choices"]) > 0:
                result_text = response_data["choices"][0]["message"]["content"]

                # JSON 부분 추출
                json_start = result_text.find("{")
                json_end = result_text.rfind("}") + 1

                if json_start >= 0 and json_end > json_start:
                    json_str = result_text[json_start:json_end]
                    try:
                        result_data = json.loads(json_str)
                        logging.info("OpenAI API를 통한 데이터 추출 성공")
                        return result_data
                    except json.JSONDecodeError as e:
                        logging.error(f"JSON 파싱 오류: {str(e)}")
                        logging.error(f"원본 JSON 문자열: {json_str}")
                else:
                    logging.error("응답에서 JSON 형식을 찾을 수 없습니다.")
                    logging.error(f"전체 응답: {result_text}")
            else:
                logging.error("응답에 choices 필드가 없습니다.")
                logging.error(f"전체 응답: {response_data}")
        else:
            logging.error(
                f"OpenAI API 호출 실패: {response.status_code} {response.reason}"
            )
            logging.error(f"응답 내용: {response.text}")

        return None

    except Exception as e:
        logging.error(f"OpenAI API 호출 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())
        return None


def export_to_structured_excel(analyzed_data, output_path):
    """분석된 데이터를 구조화된 Excel 파일로 내보내기"""
    if not analyzed_data:
        logging.error("내보낼 분석 데이터가 없습니다.")
        return False

    try:
        # 기존 파일 존재 여부 확인
        file_exists = os.path.isfile(output_path)

        if file_exists:
            # 기존 파일이 있는 경우 로드
            logging.info(f"기존 Excel 파일 '{output_path}'에 데이터를 추가합니다.")
            try:
                wb = openpyxl.load_workbook(output_path)
                # 시트가 있는지 확인
                if "추출 데이터" in wb.sheetnames:
                    ws = wb["추출 데이터"]
                    # 마지막 행 번호 찾기
                    last_row = ws.max_row
                    # 새 데이터는 마지막 행 다음에 추가
                    data_row = last_row + 1
                else:
                    # 시트가 없는 경우 새로 생성
                    ws = wb.create_sheet("추출 데이터")
                    data_row = 2  # 헤더 다음 행
            except Exception as e:
                logging.error(f"기존 Excel 파일을 열 수 없습니다: {str(e)}")
                # 파일이 손상되었거나 열 수 없는 경우 새로 생성
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "추출 데이터"
                data_row = 2  # 헤더 다음 행
                file_exists = False
        else:
            # 새 파일 생성
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "추출 데이터"
            data_row = 2  # 헤더 다음 행

        # 금액 데이터 헤더 및 값 추가
        col_idx = 1
        headers_added = False

        # 파일 정보 추가 (가장 먼저 표시)
        if "파일_정보" in analyzed_data:
            for item in analyzed_data["파일_정보"]:
                header = item.get("항목", "")

                # 헤더가 없는 경우에만 추가 (새 파일이거나 시트가 새로 생성된 경우)
                if not file_exists or not headers_added:
                    # 헤더 설정 (항목을 헤더로)
                    cell = ws.cell(row=1, column=col_idx)
                    cell.value = header
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
                    )

                # 값 추가
                cell = ws.cell(row=data_row, column=col_idx)
                cell.value = item.get("값", "")

                col_idx += 1

            # 구분선 추가 (빈 열)
            col_idx += 1
            headers_added = True

        # 금액 데이터 추가
        if "금액_데이터" in analyzed_data:
            for item in analyzed_data["금액_데이터"]:
                header = item.get("항목", "")

                # 헤더가 없는 경우에만 추가 (새 파일이거나 시트가 새로 생성된 경우)
                if not file_exists or not headers_added:
                    # 헤더 설정 (항목을 헤더로)
                    cell = ws.cell(row=1, column=col_idx)
                    cell.value = header
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
                    )

                # 값 추가
                cell = ws.cell(row=data_row, column=col_idx)
                value = item.get("금액", "")
                try:
                    # 쉼표 및 기타 문자 제거 후 숫자로 변환 시도
                    cleaned_value = value.replace(",", "").replace(" ", "")
                    if cleaned_value.isdigit():
                        cell.value = int(cleaned_value)
                        cell.number_format = "#,##0"
                    else:
                        cell.value = value
                except:
                    cell.value = value

                col_idx += 1

            headers_added = True

        # 구분선 추가 (빈 열)
        col_idx += 1

        # 은행 정보 추가
        if "은행_정보" in analyzed_data:
            for item in analyzed_data["은행_정보"]:
                header = item.get("항목", "")

                # 헤더가 없는 경우에만 추가 (새 파일이거나 시트가 새로 생성된 경우)
                if not file_exists or not headers_added:
                    # 헤더 설정 (항목을 헤더로)
                    cell = ws.cell(row=1, column=col_idx)
                    cell.value = header
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
                    )

                # 값 추가
                cell = ws.cell(row=data_row, column=col_idx)
                cell.value = item.get("값", "")

                col_idx += 1

        # 열 너비 자동 조정
        for column in ws.columns:
            max_length = 0
            column_letter = openpyxl.utils.get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width

        # 표 형식 지정 - 새로 추가된 행만 스타일 적용
        table_range = (
            f"A{data_row}:{openpyxl.utils.get_column_letter(col_idx-1)}{data_row}"
        )
        for row in ws[table_range]:
            for cell in row:
                cell.border = openpyxl.styles.Border(
                    left=openpyxl.styles.Side(style="thin"),
                    right=openpyxl.styles.Side(style="thin"),
                    top=openpyxl.styles.Side(style="thin"),
                    bottom=openpyxl.styles.Side(style="thin"),
                )

        # 파일 저장
        wb.save(output_path)
        logging.info(f"Excel 파일 저장 완료: {output_path}")
        return True

    except Exception as e:
        logging.error(f"Excel 파일 생성 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())
        return False


class FileHandler(FileSystemEventHandler):
    """파일 시스템 이벤트 처리 클래스"""

    def on_created(self, event):
        """파일 생성 이벤트 처리"""
        if event.is_directory:
            return

        file_path = event.src_path
        logging.info(f"새 파일 감지: {file_path}")

        # 파일 처리
        try:
            process_file(file_path)
        except Exception as e:
            logging.error(f"파일 처리 중 오류 발생: {str(e)}")
            logging.error(traceback.format_exc())


def is_valid_file(file_path):
    """파일 유효성 검사"""
    # 파일 경로를 Path 객체로 변환
    file_path = Path(file_path)

    # 파일 존재 확인
    if not file_path.exists():
        logging.error(f"파일이 존재하지 않습니다: {file_path}")
        return False

    # 파일 확장자 확인
    file_ext = file_path.suffix.lower()
    if file_ext not in ALL_SUPPORTED_EXTENSIONS:
        logging.error(f"지원되지 않는 파일 형식: {file_ext}")
        logging.error(
            f"현재 버전에서는 이미지 파일({', '.join(SUPPORTED_IMAGE_EXTENSIONS)}) 및 PDF 파일({SUPPORTED_PDF_EXTENSION})만 지원합니다."
        )
        return False

    # 파일 크기 확인
    file_size = file_path.stat().st_size
    if file_size > MAX_FILE_SIZE:
        logging.error(
            f"파일 크기가 너무 큽니다: {file_size} 바이트 (최대: {MAX_FILE_SIZE} 바이트)"
        )
        return False

    # PDF가 아닌 경우 MIME 타입 확인
    if file_ext != SUPPORTED_PDF_EXTENSION:
        mime_type, _ = mimetypes.guess_type(str(file_path))
        if not mime_type or not mime_type.startswith("image/"):
            logging.error(f"지원되지 않는 MIME 타입: {mime_type}")
            return False

    return True


def convert_pdf_to_images(pdf_path):
    """PDF 파일을 이미지로 변환"""
    pdf_path = Path(pdf_path)
    logging.info(f"PDF 변환 시작: {pdf_path}")

    # 날짜 기반 임시 폴더 생성
    current_date = datetime.now().strftime("%Y%m%d")
    process_id = os.getpid()
    temp_dir = Path(TEMP_DIR) / current_date / f"{pdf_path.stem}_{process_id}"

    try:
        # 임시 폴더가 이미 존재하는 경우 삭제 시도
        if temp_dir.exists():
            try:
                import shutil

                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception as e:
                logging.warning(f"기존 임시 폴더 삭제 실패: {str(e)}")
                # 폴더 삭제 실패 시 새로운 이름으로 시도
                temp_dir = (
                    Path(TEMP_DIR)
                    / current_date
                    / f"{pdf_path.stem}_{process_id}_{int(time.time())}"
                )

        # 임시 폴더 생성 (상위 폴더가 없는 경우 모두 생성)
        temp_dir.parent.mkdir(exist_ok=True, parents=True)
        temp_dir.mkdir(exist_ok=True, parents=True)

        # PyMuPDF를 사용하여 PDF를 이미지로 변환
        pdf_document = fitz.open(pdf_path)
        image_paths = []

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)

            # 고해상도 이미지로 렌더링 (300 DPI)
            pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))

            # 이미지 저장 (페이지 번호만 사용한 간단한 파일명)
            image_filename = f"page_{page_num + 1}.png"
            image_path = temp_dir / image_filename

            # 파일이 이미 존재하는 경우 삭제 시도
            if image_path.exists():
                try:
                    image_path.unlink()
                except Exception as e:
                    logging.warning(f"기존 이미지 파일 삭제 실패: {str(e)}")
                    # 삭제 실패 시 새로운 이름으로 시도
                    image_filename = f"page_{page_num + 1}_{int(time.time())}.png"
                    image_path = temp_dir / image_filename

            try:
                # 이미지 저장
                pix.save(str(image_path))
                image_paths.append(image_path)
            except Exception as e:
                logging.error(f"이미지 저장 실패 (페이지 {page_num + 1}): {str(e)}")
                # 메모리에서 이미지 데이터를 가져와 PIL로 저장 시도
                try:
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    img.save(str(image_path))
                    image_paths.append(image_path)
                    logging.info(f"대체 방법으로 이미지 저장 성공: {image_path}")
                except Exception as e2:
                    logging.error(f"대체 방법으로도 이미지 저장 실패: {str(e2)}")

        pdf_document.close()

        if not image_paths:
            logging.error(f"PDF 변환 실패: 이미지가 생성되지 않았습니다.")
            return []

        logging.info(f"PDF 변환 완료: {len(image_paths)}개 이미지 생성됨")
        return image_paths

    except Exception as e:
        logging.error(f"PDF 변환 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())
        return []


def call_clova_ocr_api(file_path):
    """네이버 클로바 OCR API 호출"""
    if not CLOVA_OCR_API_URL or not CLOVA_OCR_API_SECRET:
        logging.error(
            "CLOVA_OCR_APIGW_INVOKE_URL 또는 CLOVA_OCR_SECRET_KEY가 설정되지 않았습니다."
        )
        return None

    logging.info(f"OCR API 호출 시작: {file_path}")

    # 파일 경로를 Path 객체로 변환
    file_path = Path(file_path)

    # 파일 확장자 추출
    file_ext = file_path.suffix[1:].lower()

    # API 요청 헤더
    headers = {"X-OCR-SECRET": CLOVA_OCR_API_SECRET}

    # 파일 데이터 읽기
    with open(file_path, "rb") as f:
        file_data = f.read()

    # 커스텀 API 요청 형식 확인
    is_custom_api = "custom" in CLOVA_OCR_API_URL

    # 재시도 횟수
    max_retries = 3
    retry_count = 0

    while retry_count < max_retries:
        try:
            if is_custom_api:
                # 커스텀 OCR API 요청 형식 (multipart/form-data)
                logging.info("커스텀 OCR API 요청 형식 사용")

                # 현재 시간 타임스탬프
                timestamp = int(time.time() * 1000)

                # message JSON 생성
                message = {
                    "version": "V2",
                    "requestId": f"request_{timestamp}",
                    "timestamp": timestamp,
                    "images": [{"format": file_ext, "name": file_path.name}],
                }

                # multipart/form-data 요청 준비
                files = {
                    "file": (file_path.name, file_data, f"image/{file_ext}"),
                    "message": (None, json.dumps(message)),
                }

                response = requests.post(
                    CLOVA_OCR_API_URL, headers=headers, files=files
                )

                # 응답 확인
                if response.status_code == 200:
                    result = response.json()
                    logging.info(f"OCR API 호출 성공: {file_path}")
                    return result
                else:
                    logging.warning(
                        f"API 호출 실패 (시도 {retry_count+1}/{max_retries}): {response.status_code} {response.reason}"
                    )
                    logging.warning(f"응답 내용: {response.text}")

                    # JSON 형식으로 두 번째 시도
                    logging.info("JSON 형식으로 재시도합니다...")

                    # JSON 요청 데이터 준비
                    json_data = {
                        "version": "V2",
                        "requestId": f"request_{timestamp}",
                        "timestamp": timestamp,
                        "images": [
                            {
                                "format": file_ext.upper(),
                                "name": file_path.name,
                                "data": base64.b64encode(file_data).decode("utf-8"),
                            }
                        ],
                    }

                    # JSON 요청 헤더
                    json_headers = {
                        "X-OCR-SECRET": CLOVA_OCR_API_SECRET,
                        "Content-Type": "application/json",
                    }

                    # JSON 요청 전송
                    json_response = requests.post(
                        CLOVA_OCR_API_URL, headers=json_headers, json=json_data
                    )

                    if json_response.status_code == 200:
                        result = json_response.json()
                        logging.info(f"JSON 형식 OCR API 호출 성공: {file_path}")
                        return result
                    else:
                        logging.warning(
                            f"JSON 형식 API 호출도 실패: {json_response.status_code} {json_response.reason}"
                        )
                        logging.warning(f"응답 내용: {json_response.text}")

                        # 재시도
                        for i in range(2):  # 2번 더 시도
                            try:
                                time.sleep(2)  # 재시도 전 대기
                                response = requests.post(
                                    CLOVA_OCR_API_URL,
                                    headers=json_headers,
                                    json=json_data,
                                )
                                response.raise_for_status()
                                result = response.json()
                                logging.info(f"OCR API 재시도 성공: {file_path}")
                                return result
                            except requests.exceptions.RequestException:
                                continue

                        # 모든 재시도 실패
                        logging.error(
                            f"OCR API 호출 실패 (최대 재시도 횟수 초과): {file_path}"
                        )
            else:
                # 일반 OCR API 요청 형식 (JSON)
                logging.info("일반 OCR API 요청 형식 사용")

                # 요청 데이터 구성
                payload = {
                    "version": "V2",
                    "requestId": f"request_{int(time.time() * 1000)}",
                    "timestamp": int(time.time() * 1000),
                    "images": [
                        {
                            "format": file_ext.upper(),
                            "name": file_path.name,
                            "data": base64.b64encode(file_data).decode("utf-8"),
                        }
                    ],
                }

                # JSON 요청 헤더
                headers["Content-Type"] = "application/json"

                # 요청 전송
                response = requests.post(
                    CLOVA_OCR_API_URL, headers=headers, json=payload
                )

                # 응답 확인
                if response.status_code == 200:
                    result = response.json()
                    logging.info(f"OCR API 호출 성공: {file_path}")
                    return result
                else:
                    logging.warning(
                        f"API 호출 실패 (시도 {retry_count+1}/{max_retries}): {response.status_code} {response.reason}"
                    )
                    logging.warning(f"응답 내용: {response.text}")

                    # 재시도
                    for i in range(2):  # 2번 더 시도
                        try:
                            time.sleep(2)  # 재시도 전 대기
                            response = requests.post(
                                CLOVA_OCR_API_URL, headers=headers, json=payload
                            )
                            response.raise_for_status()
                            result = response.json()
                            logging.info(f"OCR API 재시도 성공: {file_path}")
                            return result
                        except requests.exceptions.RequestException:
                            continue

                    # 모든 재시도 실패
                    logging.error(
                        f"OCR API 호출 실패 (최대 재시도 횟수 초과): {file_path}"
                    )

            # 재시도 카운트 증가
            retry_count += 1

            # 재시도 전 대기
            if retry_count < max_retries:
                wait_time = 2**retry_count  # 지수 백오프
                logging.info(f"{wait_time}초 후 재시도합니다...")
                time.sleep(wait_time)

        except requests.exceptions.RequestException as e:
            logging.error(f"네트워크 오류 발생: {str(e)}")

            # 재시도 카운트 증가
            retry_count += 1

            # 재시도 전 대기
            if retry_count < max_retries:
                wait_time = 2**retry_count  # 지수 백오프
                logging.info(f"{wait_time}초 후 재시도합니다...")
                time.sleep(wait_time)

        except Exception as e:
            logging.error(f"예외 발생: {str(e)}")
            logging.error(traceback.format_exc())
            return None

    logging.error(f"최대 재시도 횟수({max_retries})를 초과했습니다.")
    return None


def parse_ocr_result(result):
    """OCR 결과 파싱 (원본 좌표 정보 포함)"""
    parsed_data = {"text": [], "tables": []}

    if not result or "images" not in result or not result["images"]:
        logging.error("OCR 결과가 없거나 형식이 올바르지 않습니다.")
        return parsed_data

    image_result = result["images"][0]

    # 텍스트 필드 추출 (원본 좌표 정보 포함)
    if "fields" in image_result:
        for field in image_result["fields"]:
            if "inferText" in field:
                # 텍스트 정보와 함께 원본 필드 정보도 저장
                parsed_data["text"].append(
                    {
                        "text": field["inferText"],
                        "confidence": field.get("inferConfidence", 0),
                        "field_info": field,  # 원본 필드 정보 (boundingPoly 등 포함)
                    }
                )

    # 표 데이터 추출
    if "tables" in image_result:
        for table_idx, table in enumerate(image_result["tables"]):
            table_data = {
                "table_idx": table_idx,
                "cells": [],
                "table_info": table,  # 원본 테이블 정보 저장
            }

            if "cells" in table:
                for cell in table["cells"]:
                    cell_text = ""

                    # 셀 텍스트 라인 추출
                    if "cellTextLines" in cell:
                        for line in cell["cellTextLines"]:
                            line_text = []

                            if "cellWords" in line:
                                for word in line["cellWords"]:
                                    if "inferText" in word:
                                        line_text.append(word["inferText"])

                            if line_text:
                                cell_text += " ".join(line_text) + "\n"

                    cell_text = cell_text.strip()

                    # 셀 정보와 함께 원본 셀 정보도 저장
                    table_data["cells"].append(
                        {
                            "row": cell.get("rowIndex", 0),
                            "col": cell.get("columnIndex", 0),
                            "row_span": cell.get("rowSpan", 1),
                            "col_span": cell.get("columnSpan", 1),
                            "text": cell_text,
                            "cell_info": cell,  # 원본 셀 정보 저장
                        }
                    )

            parsed_data["tables"].append(table_data)

    return parsed_data


def is_number(text):
    """텍스트가 숫자인지 확인"""
    # 쉼표와 공백 제거
    cleaned_text = text.replace(",", "").replace(" ", "")

    # 원화 기호(₩) 또는 달러 기호($) 제거
    if cleaned_text.startswith("₩") or cleaned_text.startswith("$"):
        cleaned_text = cleaned_text[1:]

    # 숫자 단위 제거 (만, 억, 천, k, m 등)
    units = ["만", "억", "천", "k", "m", "K", "M"]
    for unit in units:
        if cleaned_text.endswith(unit):
            cleaned_text = cleaned_text[: -len(unit)]

    # 소수점이 있는 경우
    if "." in cleaned_text:
        parts = cleaned_text.split(".")
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return True

    # 정수인 경우
    return cleaned_text.isdigit()


def convert_to_number(text):
    """텍스트를 숫자로 변환 (가능한 경우)"""
    try:
        # 쉼표와 공백 제거
        cleaned_text = text.replace(",", "").replace(" ", "")

        # 원화 기호(₩) 또는 달러 기호($) 제거
        if cleaned_text.startswith("₩") or cleaned_text.startswith("$"):
            cleaned_text = cleaned_text[1:]

        # 숫자 단위 처리 (만, 억, 천, k, m 등)
        multiplier = 1
        if cleaned_text.endswith("만"):
            multiplier = 10000
            cleaned_text = cleaned_text[:-1]
        elif cleaned_text.endswith("억"):
            multiplier = 100000000
            cleaned_text = cleaned_text[:-1]
        elif cleaned_text.endswith("천"):
            multiplier = 1000
            cleaned_text = cleaned_text[:-1]
        elif cleaned_text.endswith("k") or cleaned_text.endswith("K"):
            multiplier = 1000
            cleaned_text = cleaned_text[:-1]
        elif cleaned_text.endswith("m") or cleaned_text.endswith("M"):
            multiplier = 1000000
            cleaned_text = cleaned_text[:-1]

        # 숫자로 변환
        if "." in cleaned_text:
            return float(cleaned_text) * multiplier
        else:
            return int(cleaned_text) * multiplier
    except:
        return text  # 변환 실패 시 원본 텍스트 반환


def export_to_excel(parsed_data, output_path):
    """파싱된 데이터를 Excel 파일로 내보내기 (라인 기준으로 정리하고 열별로 세분화)"""
    wb = openpyxl.Workbook()

    # 기본 시트 가져오기
    main_sheet = wb.active
    main_sheet.title = "OCR 결과"

    # 텍스트 필드 정보 추출 및 정렬
    text_fields = []
    for idx, text_item in enumerate(parsed_data["text"], 1):
        # 원본 필드 정보 가져오기 (boundingPoly 정보 포함)
        field_info = text_item.get("field_info", {})

        # 좌표 정보 추출
        vertices = field_info.get("boundingPoly", {}).get("vertices", [])
        if vertices:
            # 좌표의 평균값 계산
            x_coords = [vertex.get("x", 0) for vertex in vertices]
            y_coords = [vertex.get("y", 0) for vertex in vertices]
            avg_x = sum(x_coords) / len(x_coords) if x_coords else 0
            avg_y = sum(y_coords) / len(y_coords) if y_coords else 0

            # 텍스트 정보 저장
            text_fields.append(
                {
                    "idx": idx,
                    "text": text_item["text"],
                    "confidence": text_item["confidence"],
                    "x": avg_x,
                    "y": avg_y,
                    "width": max(x_coords) - min(x_coords) if x_coords else 0,
                    "height": max(y_coords) - min(y_coords) if y_coords else 0,
                }
            )

    # Y좌표로 정렬 (위에서 아래로)
    text_fields.sort(key=lambda x: x["y"])

    # 라인 그룹화 (Y좌표가 비슷한 텍스트를 같은 라인으로 그룹화)
    line_threshold = 20  # Y좌표 차이가 이 값보다 작으면 같은 라인으로 간주
    lines = []
    current_line = []

    for i, field in enumerate(text_fields):
        if i == 0:
            # 첫 번째 필드는 현재 라인에 추가
            current_line.append(field)
        else:
            # 이전 필드와 Y좌표 차이 계산
            prev_y = text_fields[i - 1]["y"]
            curr_y = field["y"]

            if abs(curr_y - prev_y) <= line_threshold:
                # 같은 라인으로 간주
                current_line.append(field)
            else:
                # 새로운 라인 시작
                if current_line:
                    # 현재 라인을 X좌표로 정렬
                    current_line.sort(key=lambda x: x["x"])
                    lines.append(current_line)
                current_line = [field]

    # 마지막 라인 추가
    if current_line:
        current_line.sort(key=lambda x: x["x"])
        lines.append(current_line)

    # 열 구분을 위한 X좌표 클러스터링
    all_x_coords = [field["x"] for field in text_fields]

    # X좌표 클러스터링 (K-means 대신 간단한 방법 사용)
    def cluster_x_coordinates(x_coords, threshold=100):
        if not x_coords:
            return []

        # 정렬된 X좌표
        sorted_x = sorted(x_coords)

        # 클러스터 초기화
        clusters = [[sorted_x[0]]]

        # 각 X좌표를 적절한 클러스터에 할당
        for x in sorted_x[1:]:
            # 이전 클러스터의 마지막 값과 비교
            if x - clusters[-1][-1] > threshold:
                # 새 클러스터 시작
                clusters.append([x])
            else:
                # 기존 클러스터에 추가
                clusters[-1].append(x)

        # 각 클러스터의 중심값 계산
        cluster_centers = [sum(cluster) / len(cluster) for cluster in clusters]
        return cluster_centers

    # X좌표 클러스터 중심 계산
    x_clusters = cluster_x_coordinates(all_x_coords)
    max_columns = len(x_clusters) + 1  # 여유 있게 열 개수 설정

    # 헤더 추가 (열 개수에 맞게)
    headers = ["라인"] + [f"열_{i+1}" for i in range(max_columns)] + ["신뢰도"]
    main_sheet.append(headers)

    # 라인별로 엑셀에 추가
    for line_idx, line in enumerate(lines, 1):
        # 라인의 평균 신뢰도 계산
        avg_confidence = (
            sum([field["confidence"] for field in line]) / len(line) if line else 0
        )

        # 각 필드를 적절한 열에 배치
        row_data = [line_idx] + [""] * max_columns + [f"{avg_confidence:.4f}"]

        for field in line:
            # 가장 가까운 클러스터 찾기
            if x_clusters:
                closest_cluster_idx = min(
                    range(len(x_clusters)),
                    key=lambda i: abs(field["x"] - x_clusters[i]),
                )
                # 해당 열에 텍스트 추가 (이미 텍스트가 있으면 공백으로 구분하여 추가)
                col_idx = closest_cluster_idx + 1  # 첫 번째 열은 라인 번호
                if row_data[col_idx] == "":
                    row_data[col_idx] = field["text"]
                else:
                    row_data[col_idx] += " " + field["text"]
            else:
                # 클러스터가 없는 경우 첫 번째 데이터 열에 추가
                if row_data[1] == "":
                    row_data[1] = field["text"]
                else:
                    row_data[1] += " " + field["text"]

        # 엑셀에 라인 추가
        main_sheet.append(row_data)

    # 표 시트 생성
    if parsed_data["tables"]:
        for table_idx, table_data in enumerate(parsed_data["tables"], 1):
            table_sheet = wb.create_sheet(f"표_{table_idx}")

            # 표 데이터 구성
            if table_data["cells"]:
                # 행과 열 크기 결정
                max_row = max([cell["row"] for cell in table_data["cells"]]) + 1
                max_col = max([cell["col"] for cell in table_data["cells"]]) + 1

                # 2D 배열 초기화
                table_array = [[None for _ in range(max_col)] for _ in range(max_row)]

                # 셀 데이터 채우기
                for cell in table_data["cells"]:
                    row = cell["row"]
                    col = cell["col"]
                    cell_text = cell["text"]

                    # 숫자 데이터 처리
                    if is_number(cell_text):
                        # 숫자로 변환
                        cell_value = convert_to_number(cell_text)
                    else:
                        cell_value = cell_text

                    table_array[row][col] = cell_value

                # 시트에 데이터 추가
                for row_idx, row_data in enumerate(table_array, 1):
                    # None 값을 빈 문자열로 변환
                    formatted_row = ["" if cell is None else cell for cell in row_data]

                    # 행 추가
                    for col_idx, cell_value in enumerate(formatted_row, 1):
                        cell = table_sheet.cell(row=row_idx, column=col_idx)
                        cell.value = cell_value

                        # 숫자 형식 지정
                        if isinstance(cell_value, (int, float)):
                            # 금액 형식으로 표시
                            cell.number_format = "#,##0"

                # 표 서식 지정
                for row_idx in range(1, max_row + 1):
                    for col_idx in range(1, max_col + 1):
                        cell = table_sheet.cell(row=row_idx, column=col_idx)
                        # 테두리 설정
                        cell.border = openpyxl.styles.Border(
                            left=openpyxl.styles.Side(style="thin"),
                            right=openpyxl.styles.Side(style="thin"),
                            top=openpyxl.styles.Side(style="thin"),
                            bottom=openpyxl.styles.Side(style="thin"),
                        )

                # 첫 번째 행을 헤더로 설정
                for col_idx in range(1, max_col + 1):
                    cell = table_sheet.cell(row=1, column=col_idx)
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
                    )

                # 자동 필터 설정
                table_sheet.auto_filter.ref = (
                    f"A1:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
                )

                # 열 너비 자동 조정
                for col_idx in range(1, max_col + 1):
                    column_letter = openpyxl.utils.get_column_letter(col_idx)
                    max_length = 0
                    for row_idx in range(1, max_row + 1):
                        cell = table_sheet.cell(row=row_idx, column=col_idx)
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    adjusted_width = (max_length + 2) * 1.2
                    table_sheet.column_dimensions[column_letter].width = adjusted_width

    # 원본 데이터 시트 추가 (디버깅용)
    debug_sheet = wb.create_sheet("원본 데이터")
    debug_sheet.append(
        ["번호", "텍스트", "신뢰도", "X좌표", "Y좌표", "너비", "높이", "숫자여부"]
    )

    for idx, field in enumerate(text_fields, 1):
        debug_sheet.append(
            [
                idx,
                field["text"],
                field["confidence"],
                field["x"],
                field["y"],
                field["width"],
                field["height"],
                "숫자" if is_number(field["text"]) else "텍스트",
            ]
        )

    # 클러스터 정보 시트 추가 (디버깅용)
    cluster_sheet = wb.create_sheet("X좌표 클러스터")
    cluster_sheet.append(["클러스터 번호", "중심 X좌표"])

    for idx, center in enumerate(x_clusters, 1):
        cluster_sheet.append([idx, center])

    # 열 너비 자동 조정
    for sheet in wb.worksheets:
        for column in sheet.columns:
            max_length = 0
            column_letter = openpyxl.utils.get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
            adjusted_width = (max_length + 2) * 1.2
            sheet.column_dimensions[column_letter].width = adjusted_width

    # 파일 저장
    wb.save(output_path)
    logging.info(f"Excel 파일 저장 완료: {output_path}")
    logging.info(
        f"총 {len(lines)}개 라인, {len(x_clusters)}개 열, {len(parsed_data['tables'])}개 표 저장됨"
    )


def merge_ocr_results(results):
    """여러 OCR 결과를 하나로 병합"""
    merged_data = {"text": [], "tables": []}

    table_count = 0

    for result in results:
        # 텍스트 병합
        merged_data["text"].extend(result["text"])

        # 표 병합 (테이블 인덱스 조정)
        for table in result["tables"]:
            table_copy = table.copy()
            table_copy["table_idx"] = table_count
            merged_data["tables"].append(table_copy)
            table_count += 1

    return merged_data


def process_file(file_path):
    """파일 처리 함수"""
    logging.info(f"파일 처리 시작: {file_path}")

    # 파일 경로를 Path 객체로 변환
    file_path = Path(file_path)

    # 파일 유효성 검사
    if not is_valid_file(file_path):
        logging.error(f"파일 유효성 검사 실패: {file_path}")
        return

    try:
        # 결과 파일 경로 생성
        file_name = file_path.name
        file_base_name = file_path.stem
        current_date = datetime.now().strftime("%Y%m%d")
        process_id = os.getpid()  # 프로세스 ID 추가

        # 결과 디렉토리 확인 및 생성 (날짜별 폴더 구조)
        result_dir = Path(RESULT_DIR) / current_date
        if not result_dir.exists():
            try:
                result_dir.mkdir(parents=True, exist_ok=True)
                logging.info(f"결과 디렉토리 생성: {result_dir}")
            except Exception as e:
                logging.error(f"결과 디렉토리 생성 실패: {str(e)}")
                # 대체 경로 사용
                result_dir = Path(os.getcwd()) / "results" / current_date
                result_dir.mkdir(parents=True, exist_ok=True)
                logging.info(f"대체 결과 디렉토리 생성: {result_dir}")

        # 공통 output.xlsx 파일 경로 설정
        common_output_path = result_dir / "output.xlsx"

        # PDF 파일 처리
        if file_path.suffix.lower() == SUPPORTED_PDF_EXTENSION:
            logging.info(f"PDF 파일 처리 시작: {file_path}")

            # PDF를 이미지로 변환
            image_paths = convert_pdf_to_images(file_path)

            if not image_paths:
                logging.error(
                    f"PDF 변환 실패 또는 변환된 이미지가 없습니다: {file_path}"
                )
                return

            # 각 이미지에 대해 OCR 처리
            all_ocr_results = []

            for i, image_path in enumerate(image_paths):
                logging.info(
                    f"PDF 페이지 {i+1}/{len(image_paths)} 처리 중: {image_path}"
                )

                # OCR API 호출
                ocr_result = call_clova_ocr_api(image_path)

                if not ocr_result:
                    logging.error(f"OCR API 호출 실패: {image_path}")
                    continue

                # JSON 결과 저장 (각 페이지별)
                page_json_path = (
                    result_dir / f"{file_base_name}_page{i+1}_{process_id}.json"
                )

                try:
                    with open(page_json_path, "w", encoding="utf-8") as f:
                        json.dump(ocr_result, f, ensure_ascii=False, indent=2)
                except Exception as e:
                    logging.error(f"JSON 결과 저장 실패: {str(e)}")
                    # 대체 경로 시도
                    page_json_path = (
                        Path(os.getcwd())
                        / f"{file_base_name}_page{i+1}_{process_id}.json"
                    )
                    with open(page_json_path, "w", encoding="utf-8") as f:
                        json.dump(ocr_result, f, ensure_ascii=False, indent=2)

                all_ocr_results.append(ocr_result)

            # 모든 OCR 결과 분석
            for i, ocr_result in enumerate(all_ocr_results):
                # OpenAI API를 사용하여 OCR 결과 분석
                analyzed_data = analyze_ocr_with_openai(ocr_result)

                if analyzed_data:
                    # 원본 파일 정보 추가
                    if "파일_정보" not in analyzed_data:
                        analyzed_data["파일_정보"] = []

                    analyzed_data["파일_정보"].append(
                        {"항목": "원본파일명", "값": file_path.name}
                    )

                    analyzed_data["파일_정보"].append(
                        {"항목": "처리시간", "값": current_date}
                    )

                    analyzed_data["파일_정보"].append(
                        {"항목": "페이지번호", "값": f"{i+1}/{len(all_ocr_results)}"}
                    )

                    # 구조화된 Excel 파일로 내보내기 (공통 파일에 추가)
                    try:
                        export_to_structured_excel(analyzed_data, common_output_path)
                        logging.info(
                            f"데이터가 공통 Excel 파일에 추가되었습니다: {common_output_path}"
                        )
                    except Exception as e:
                        logging.error(f"공통 Excel 파일 저장 실패: {str(e)}")
                        # 대체 경로 시도
                        alt_output_path = Path(os.getcwd()) / "output.xlsx"
                        try:
                            export_to_structured_excel(analyzed_data, alt_output_path)
                            logging.info(
                                f"대체 경로에 Excel 파일 저장 성공: {alt_output_path}"
                            )
                        except Exception as e2:
                            logging.error(
                                f"대체 경로에도 Excel 파일 저장 실패: {str(e2)}"
                            )
                            # 마지막 대안으로 개별 파일 저장
                            individual_output_path = (
                                result_dir
                                / f"output_{file_base_name}_{current_date}.xlsx"
                            )
                            export_to_structured_excel(
                                analyzed_data, individual_output_path
                            )
                            logging.info(
                                f"개별 Excel 파일 저장 성공: {individual_output_path}"
                            )

            # 처리 완료 후 임시 파일 정리 시도
            try:
                for image_path in image_paths:
                    if image_path.exists():
                        try:
                            image_path.unlink()
                        except:
                            pass  # 삭제 실패 무시

                # 임시 디렉토리 삭제 시도
                temp_dir = image_paths[0].parent if image_paths else None
                if temp_dir and temp_dir.exists():
                    try:
                        import shutil

                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except:
                        pass  # 삭제 실패 무시
            except Exception as e:
                logging.warning(f"임시 파일 정리 중 오류 발생 (무시됨): {str(e)}")

            logging.info(f"PDF 파일 처리 완료: {file_path}")
        else:
            # 이미지 파일 처리
            logging.info(f"이미지 파일 처리 시작: {file_path}")

            # OCR API 호출
            ocr_result = call_clova_ocr_api(file_path)

            if not ocr_result:
                logging.error(f"OCR API 호출 실패: {file_path}")
                return

            # JSON 결과 저장
            json_result_path = result_dir / f"{file_base_name}_{process_id}.json"
            with open(json_result_path, "w", encoding="utf-8") as f:
                json.dump(ocr_result, f, ensure_ascii=False, indent=2)

            # OpenAI API를 사용하여 OCR 결과 분석
            analyzed_data = analyze_ocr_with_openai(ocr_result)

            if analyzed_data:
                # 원본 파일 정보 추가
                if "파일_정보" not in analyzed_data:
                    analyzed_data["파일_정보"] = []

                analyzed_data["파일_정보"].append(
                    {"항목": "원본파일명", "값": file_path.name}
                )

                analyzed_data["파일_정보"].append(
                    {"항목": "처리시간", "값": current_date}
                )

                # 구조화된 Excel 파일로 내보내기 (공통 파일에 추가)
                try:
                    export_to_structured_excel(analyzed_data, common_output_path)
                    logging.info(
                        f"데이터가 공통 Excel 파일에 추가되었습니다: {common_output_path}"
                    )
                except Exception as e:
                    logging.error(f"공통 Excel 파일 저장 실패: {str(e)}")
                    # 대체 경로 시도
                    alt_output_path = Path(os.getcwd()) / "output.xlsx"
                    try:
                        export_to_structured_excel(analyzed_data, alt_output_path)
                        logging.info(
                            f"대체 경로에 Excel 파일 저장 성공: {alt_output_path}"
                        )
                    except Exception as e2:
                        logging.error(f"대체 경로에도 Excel 파일 저장 실패: {str(e2)}")
                        # 마지막 대안으로 개별 파일 저장
                        individual_output_path = (
                            result_dir / f"output_{file_base_name}_{current_date}.xlsx"
                        )
                        export_to_structured_excel(
                            analyzed_data, individual_output_path
                        )
                        logging.info(
                            f"개별 Excel 파일 저장 성공: {individual_output_path}"
                        )

            logging.info(f"이미지 파일 처리 완료: {file_path}")

    except Exception as e:
        logging.error(f"파일 처리 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())


def process_existing_files():
    """기존 파일 처리 함수"""
    logging.info("기존 파일 처리 시작")

    # 소스 디렉토리의 모든 파일 가져오기
    source_dir = Path(SOURCE_DIR)
    files = list(source_dir.glob("*"))

    if not files:
        logging.info("처리할 파일이 없습니다.")
        return

    logging.info(f"총 {len(files)}개 파일 발견")

    # 각 파일 처리
    for file_path in files:
        if file_path.is_file():
            try:
                process_file(file_path)
            except Exception as e:
                logging.error(f"파일 처리 중 오류 발생: {str(e)}")
                logging.error(traceback.format_exc())

    logging.info("기존 파일 처리 완료")


def main():
    """메인 함수"""
    logging.info("프로그램 시작")
    logging.info(f"소스 디렉토리: {os.path.abspath(SOURCE_DIR)}")
    logging.info(f"결과 디렉토리: {os.path.abspath(RESULT_DIR)}")

    # 환경 변수 확인
    if not CLOVA_OCR_API_SECRET:
        logging.warning("경고: CLOVA_OCR_SECRET_KEY 환경 변수가 설정되지 않았습니다.")
        print("경고: CLOVA_OCR_SECRET_KEY 환경 변수를 설정해야 합니다.")
        print(
            "  .env 파일에 CLOVA_OCR_SECRET_KEY=your_api_key_here 형식으로 추가하거나"
        )
        print("  환경 변수로 직접 설정하세요.")

    if not OPENAI_API_KEY:
        logging.warning("경고: OPENAI_API_KEY 환경 변수가 설정되지 않았습니다.")
        print("경고: OPENAI_API_KEY 환경 변수를 설정해야 합니다.")
        print("  .env 파일에 OPENAI_API_KEY=your_api_key_here 형식으로 추가하거나")
        print("  환경 변수로 직접 설정하세요.")

    # API URL 확인
    logging.info(f"CLOVA OCR API URL: {CLOVA_OCR_API_URL}")

    # 기존 파일 처리
    process_existing_files()

    try:
        # 파일 시스템 감시 설정
        event_handler = FileHandler()
        observer = Observer()
        observer.schedule(event_handler, SOURCE_DIR, recursive=False)

        # 감시 시작
        try:
            observer.start()
            logging.info(
                f"{SOURCE_DIR} 디렉토리 감시 중... (종료하려면 Ctrl+C를 누르세요)"
            )

            # 무한 루프로 감시 유지
            while True:
                time.sleep(1)
        except Exception as e:
            logging.error(f"감시 시작 중 오류 발생: {str(e)}")
            logging.error(traceback.format_exc())

            # 대체 방법: 기존 파일만 처리하고 종료
            logging.info(
                "파일 감시 기능을 사용할 수 없습니다. 기존 파일만 처리하고 종료합니다."
            )

    except KeyboardInterrupt:
        if "observer" in locals() and observer.is_alive():
            observer.stop()
            observer.join()

    except Exception as e:
        logging.error(f"프로그램 실행 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())

    logging.info("프로그램 종료")


if __name__ == "__main__":
    main()
