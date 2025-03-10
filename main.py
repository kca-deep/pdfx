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
from watchdog.observers.polling import PollingObserver as Observer  # Python 3.13 호환성을 위해 PollingObserver 사용
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv
import openpyxl
import fitz  # PyMuPDF
from PIL import Image
import io
import traceback
from tqdm import tqdm  # 진행률 표시를 위한 라이브러리

# 환경 변수 로드
load_dotenv(override=True)  # override=True로 설정하여 기존 환경 변수를 덮어씁니다.

# 로그 디렉토리 확인 및 생성
log_dir = Path("logs")
if not log_dir.exists():
    log_dir.mkdir(parents=True, exist_ok=True)

# 날짜 기반 로그 파일명 생성 (하루 단위)
current_date = datetime.now().strftime("%Y%m%d")
log_filename = log_dir / f"pdfx_{current_date}.log"

# 로깅 설정 최적화 - 콘솔 출력 제거, 파일에만 기록
logging.basicConfig(
    level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(str(log_filename), mode="a", encoding="utf-8"),
    ],
)

# 로그 파일 경로 기록 (파일에만 기록)
logging.info(f"로그 파일 경로: {log_filename}")

# 네이버 클로바 OCR API 설정
CLOVA_OCR_API_URL = os.getenv("CLOVA_OCR_APIGW_INVOKE_URL")
CLOVA_OCR_API_SECRET = os.getenv("CLOVA_OCR_SECRET_KEY")

# OpenAI API 설정
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_API_URL = "https://api.openai.com/v1/chat/completions"
# openai.api_key = OPENAI_API_KEY  # requests 라이브러리로 대체

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

# 콘솔 출력 설정 - 진행률 표시를 위한 tqdm 사용
print(f"PDF-X 프로그램 시작 (로그 파일: {log_filename})")
print(f"소스 디렉토리: {os.path.abspath(SOURCE_DIR)}")
print(f"결과 디렉토리: {os.path.abspath(RESULT_DIR)}")

# OpenAI API 사용량 추적 클래스
class APIUsageTracker:
    def __init__(self, exchange_rate=1450):
        self.total_tokens = 0
        self.total_prompt_tokens = 0
        self.total_completion_tokens = 0
        self.total_cost_usd = 0.0
        self.exchange_rate = exchange_rate  # 1 USD = 1450 KRW
        self.model_costs = {
            "gpt-4o-mini": {
                "input": 0.000002,  # 입력 토큰당 비용 (USD) - 1000토큰당 $0.002
                "output": 0.000002,  # 출력 토큰당 비용 (USD) - 1000토큰당 $0.002
            },
            "gpt-4o": {
                "input": 0.000003,  # 1000토큰당 $0.003
                "output": 0.000006,  # 1000토큰당 $0.006
            },
            "gpt-3.5-turbo": {
                "input": 0.0000005,  # 1000토큰당 $0.0005
                "output": 0.0000015,  # 1000토큰당 $0.0015
            }
        }
        self.default_model = "gpt-4o-mini"
    
    def count_tokens(self, text):
        """텍스트의 토큰 수를 대략적으로 계산합니다. (tiktoken 없이 간단한 방식 사용)"""
        # 영어 기준으로 단어 4개당 약 3개의 토큰으로 계산 (OpenAI 문서 기준)
        # 한글은 더 많은 토큰을 사용하므로 글자 수 기준으로 계산
        words = text.split()
        word_count = len(words)
        char_count = len(text)
        
        # 한글이 포함된 경우 글자 수 기준으로 계산 (한글 1글자당 약 1.5 토큰)
        has_korean = any(ord(char) >= 0xAC00 and ord(char) <= 0xD7A3 for char in text)
        
        if has_korean:
            return int(char_count * 1.5)
        else:
            return int(word_count * 0.75)
    
    def update_usage(self, prompt_tokens, completion_tokens, model=None):
        """API 사용량을 업데이트합니다."""
        if model is None:
            model = self.default_model
            
        if model not in self.model_costs:
            model = self.default_model
            
        # 토큰 수 업데이트
        self.total_prompt_tokens += prompt_tokens
        self.total_completion_tokens += completion_tokens
        self.total_tokens += prompt_tokens + completion_tokens
        
        # 비용 계산
        prompt_cost = prompt_tokens * self.model_costs[model]["input"]
        completion_cost = completion_tokens * self.model_costs[model]["output"]
        cost_usd = prompt_cost + completion_cost
        self.total_cost_usd += cost_usd
        
        # 로그에 사용량 기록
        cost_krw = cost_usd * self.exchange_rate
        total_cost_krw = self.total_cost_usd * self.exchange_rate
        
        # 1000토큰 단위로 비용 표시 (가독성 향상)
        prompt_cost_per_1k = self.model_costs[model]["input"] * 1000
        completion_cost_per_1k = self.model_costs[model]["output"] * 1000
        
        logging.info(f"API 사용량 업데이트: 모델={model}, 입력 토큰={prompt_tokens}, 출력 토큰={completion_tokens}")
        logging.info(f"토큰 비용 (1000토큰당): 입력=${prompt_cost_per_1k:.4f}, 출력=${completion_cost_per_1k:.4f}")
        logging.info(f"현재 호출 비용: ${cost_usd:.6f} (₩{cost_krw:.2f})")
        logging.info(f"누적 API 사용량: 총 토큰={self.total_tokens}, 입력={self.total_prompt_tokens}, 출력={self.total_completion_tokens}")
        logging.info(f"누적 API 비용: ${self.total_cost_usd:.6f} (₩{total_cost_krw:.2f})")
        
        return {
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": prompt_tokens + completion_tokens,
            "cost_usd": cost_usd,
            "cost_krw": cost_krw
        }
    
    def get_usage_summary(self):
        """API 사용량 요약을 반환합니다."""
        return {
            "total_tokens": self.total_tokens,
            "prompt_tokens": self.total_prompt_tokens,
            "completion_tokens": self.total_completion_tokens,
            "cost_usd": self.total_cost_usd,
            "cost_krw": self.total_cost_usd * self.exchange_rate
        }

# API 사용량 추적기 인스턴스 생성
api_tracker = APIUsageTracker(exchange_rate=1450)

# 설정 로그
logging.info(f"CLOVA OCR API URL: {CLOVA_OCR_API_URL}")
if CLOVA_OCR_API_SECRET:
    logging.info(
        f"CLOVA OCR API SECRET: {CLOVA_OCR_API_SECRET[:4]}...{CLOVA_OCR_API_SECRET[-4:]} (길이: {len(CLOVA_OCR_API_SECRET)})"
    )


def analyze_ocr_with_openai(ocr_result):
    """OpenAI API를 사용하여 OCR 결과 분석 (requests 라이브러리 사용)"""
    if not OPENAI_API_KEY:
        logging.error("OpenAI API 키가 설정되지 않았습니다.")
        return None

    try:
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

        # 사용할 모델
        model = "gpt-4o-mini"
        
        # 토큰 수 계산
        prompt_tokens = api_tracker.count_tokens(prompt)
        
        # OpenAI API 요청 데이터 준비
        payload = {
            "model": model,
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

        # API 요청
        response = requests.post(
            OPENAI_API_URL, headers=headers, data=json.dumps(payload)
        )

        # 응답 확인
        if response.status_code == 200:
            result = response.json()
            
            # 응답에서 토큰 사용량 추출
            usage = result.get("usage", {})
            prompt_tokens = usage.get("prompt_tokens", 0)
            completion_tokens = usage.get("completion_tokens", 0)
            
            # API 사용량 업데이트 (응답에서 실제 토큰 수를 가져옴)
            if prompt_tokens == 0:  # API가 토큰 수를 반환하지 않는 경우 추정값 사용
                prompt_tokens = api_tracker.count_tokens(prompt)
                
            api_tracker.update_usage(prompt_tokens, completion_tokens, model)
            
            # 응답 텍스트 추출
            response_text = result["choices"][0]["message"]["content"]

            try:
                # JSON 형식으로 파싱
                json_start = response_text.find("{")
                json_end = response_text.rfind("}") + 1
                if json_start >= 0 and json_end > json_start:
                    json_text = response_text[json_start:json_end]
                    return json.loads(json_text)
                else:
                    logging.error("응답에서 JSON 형식을 찾을 수 없습니다.")
                    return None
            except json.JSONDecodeError as e:
                logging.error(f"JSON 파싱 오류: {str(e)}")
                logging.error(f"원본 응답: {response_text}")
                return None
        else:
            logging.error(f"API 요청 실패: {response.status_code} - {response.text}")
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
        return False

    # 파일 크기 확인
    file_size = file_path.stat().st_size
    if file_size > MAX_FILE_SIZE:
        logging.error(f"파일 크기가 너무 큽니다: {file_size} 바이트 (최대: {MAX_FILE_SIZE} 바이트)")
        return False

    # PDF가 아닌 경우 MIME 타입 확인
    if file_ext != SUPPORTED_PDF_EXTENSION:
        mime_type, _ = mimetypes.guess_type(str(file_path))
        if not mime_type or not mime_type.startswith("image/"):
            logging.error(f"지원되지 않는 MIME 타입: {mime_type}")
            return False

    return True


def convert_pdf_to_images(pdf_path, pbar=None):
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
                temp_dir = Path(TEMP_DIR) / current_date / f"{pdf_path.stem}_{process_id}_{int(time.time())}"

        # 임시 폴더 생성 (상위 폴더가 없는 경우 모두 생성)
        temp_dir.parent.mkdir(exist_ok=True, parents=True)
        temp_dir.mkdir(exist_ok=True, parents=True)

        # PyMuPDF를 사용하여 PDF를 이미지로 변환
        pdf_document = fitz.open(pdf_path)
        image_paths = []
        
        # PDF 페이지 변환
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
                if pbar:
                    pbar.update(1)
                    pbar.refresh()
            except Exception as e:
                logging.error(f"이미지 저장 실패 (페이지 {page_num + 1}): {str(e)}")

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


def process_file(file_path, pbar=None):
    """파일 처리 함수"""
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

        # 결과 디렉토리 확인 및 생성 (날짜별 폴더 구조)
        result_dir = Path(RESULT_DIR) / current_date
        if not result_dir.exists():
            try:
                result_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                logging.error(f"결과 디렉토리 생성 실패: {str(e)}")
                # 대체 경로 사용
                result_dir = Path(os.getcwd()) / "results" / current_date
                result_dir.mkdir(parents=True, exist_ok=True)

        # 공통 output.xlsx 파일 경로 설정
        common_output_path = result_dir / "output.xlsx"

        # PDF 파일 처리
        if file_path.suffix.lower() == SUPPORTED_PDF_EXTENSION:
            if pbar:
                pbar.set_description(f"PDF 변환 중: {file_path.name}")
                pbar.refresh()

            # PDF를 이미지로 변환
            image_paths = convert_pdf_to_images(file_path, pbar)

            if not image_paths:
                logging.error(f"PDF 변환 실패 또는 변환된 이미지가 없습니다: {file_path}")
                return

            # 각 이미지에 대해 OCR 처리
            all_ocr_results = []
            
            if pbar:
                pbar.set_description(f"OCR 처리 중: {file_path.name}")
                pbar.refresh()

            for i, image_path in enumerate(image_paths):
                # OCR API 호출
                ocr_result = call_clova_ocr_api(image_path)
                
                if pbar:
                    pbar.update(1)
                    pbar.refresh()

                if not ocr_result:
                    logging.error(f"OCR API 호출 실패: {image_path}")
                    continue

                # JSON 결과 저장 (각 페이지별)
                page_json_path = result_dir / f"{file_base_name}_page{i+1}_{os.getpid()}.json"

                try:
                    with open(page_json_path, "w", encoding="utf-8") as f:
                        json.dump(ocr_result, f, ensure_ascii=False, indent=2)
                except Exception as e:
                    logging.error(f"JSON 결과 저장 실패: {str(e)}")
                    # 대체 경로 시도
                    page_json_path = Path(os.getcwd()) / f"{file_base_name}_page{i+1}_{os.getpid()}.json"
                    with open(page_json_path, "w", encoding="utf-8") as f:
                        json.dump(ocr_result, f, ensure_ascii=False, indent=2)

                all_ocr_results.append(ocr_result)

            # 모든 OCR 결과 분석
            if pbar:
                pbar.set_description(f"데이터 분석 중: {file_path.name}")
                pbar.refresh()

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
                        logging.info(f"데이터가 공통 Excel 파일에 추가되었습니다: {common_output_path}")
                    except Exception as e:
                        logging.error(f"공통 Excel 파일 저장 실패: {str(e)}")
                        # 대체 경로 시도
                        alt_output_path = Path(os.getcwd()) / "output.xlsx"
                        try:
                            export_to_structured_excel(analyzed_data, alt_output_path)
                            logging.info(f"대체 경로에 Excel 파일 저장 성공: {alt_output_path}")
                        except Exception as e2:
                            logging.error(f"대체 경로에도 Excel 파일 저장 실패: {str(e2)}")
                            # 마지막 대안으로 개별 파일 저장
                            individual_output_path = result_dir / f"output_{file_base_name}_{current_date}.xlsx"
                            export_to_structured_excel(analyzed_data, individual_output_path)
                            logging.info(f"개별 Excel 파일 저장 성공: {individual_output_path}")

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

            if pbar:
                pbar.set_description(f"OCR 처리 중: {file_path.name}")
                pbar.refresh()

            # OCR API 호출
            ocr_result = call_clova_ocr_api(file_path)

            if pbar:
                pbar.update(1)
                pbar.refresh()

            if not ocr_result:
                logging.error(f"OCR API 호출 실패: {file_path}")
                return

            # JSON 결과 저장
            json_result_path = result_dir / f"{file_base_name}_{os.getpid()}.json"
            with open(json_result_path, "w", encoding="utf-8") as f:
                json.dump(ocr_result, f, ensure_ascii=False, indent=2)

            if pbar:
                pbar.set_description(f"데이터 분석 중: {file_path.name}")
                pbar.refresh()

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
                    logging.info(f"데이터가 공통 Excel 파일에 추가되었습니다: {common_output_path}")
                except Exception as e:
                    logging.error(f"공통 Excel 파일 저장 실패: {str(e)}")
                    # 대체 경로 시도
                    alt_output_path = Path(os.getcwd()) / "output.xlsx"
                    try:
                        export_to_structured_excel(analyzed_data, alt_output_path)
                        logging.info(f"대체 경로에 Excel 파일 저장 성공: {alt_output_path}")
                    except Exception as e2:
                        logging.error(f"대체 경로에도 Excel 파일 저장 실패: {str(e2)}")
                        # 마지막 대안으로 개별 파일 저장
                        individual_output_path = result_dir / f"output_{file_base_name}_{current_date}.xlsx"
                        export_to_structured_excel(analyzed_data, individual_output_path)
                        logging.info(f"개별 Excel 파일 저장 성공: {individual_output_path}")

            logging.info(f"이미지 파일 처리 완료: {file_path}")

    except Exception as e:
        logging.error(f"파일 처리 중 오류 발생: {str(e)}")
        logging.error(traceback.format_exc())


def process_existing_files():
    """기존 파일 처리"""
    # 소스 디렉토리 확인
    source_dir = Path(SOURCE_DIR)
    if not source_dir.exists():
        try:
            source_dir.mkdir(parents=True, exist_ok=True)
            logging.info(f"소스 디렉토리 생성: {source_dir}")
        except Exception as e:
            logging.error(f"소스 디렉토리 생성 실패: {str(e)}")
            return

    # 지원되는 파일 확장자 목록
    supported_extensions = ALL_SUPPORTED_EXTENSIONS

    # 소스 디렉토리에서 지원되는 파일 찾기
    files = [f for f in source_dir.glob("*") if f.suffix.lower() in supported_extensions]
    
    if not files:
        return
        
    logging.info(f"총 {len(files)}개 파일 발견")
    
    # 전체 작업량 계산
    total_steps = 0
    for file_path in files:
        if file_path.suffix.lower() == SUPPORTED_PDF_EXTENSION:
            # PDF 파일인 경우 페이지 수 확인
            try:
                pdf_document = fitz.open(file_path)
                total_steps += len(pdf_document) * 2  # PDF 변환 및 OCR 처리
                pdf_document.close()
            except Exception as e:
                logging.error(f"PDF 페이지 수 확인 실패: {str(e)}")
                total_steps += 2  # 기본값으로 2단계 추가
        else:
            total_steps += 1  # 이미지 파일은 1단계

    # tqdm을 사용하여 통합 진행률 표시
    with tqdm(total=total_steps, desc="파일 처리", ncols=80, dynamic_ncols=False, file=sys.stdout) as pbar:
        for file_path in files:
            if file_path.is_file():
                try:
                    process_file(file_path, pbar)
                except Exception as e:
                    logging.error(f"파일 처리 중 오류 발생: {str(e)}")
                    logging.error(traceback.format_exc())


def main():
    """메인 함수"""
    logging.info("프로그램 시작")
    
    # 환경 변수 로드
    load_dotenv()

    # 필수 환경 변수 검증
    required_env_vars = [
        "CLOVA_OCR_APIGW_INVOKE_URL",
        "CLOVA_OCR_SECRET_KEY",
        "OPENAI_API_KEY",
    ]
    for var in required_env_vars:
        if not os.getenv(var):
            logging.error(f"필수 환경 변수 {var}가 설정되지 않았습니다.")
            sys.exit(1)

    # 환경 변수 로드
    CLOVA_OCR_API_URL = os.getenv("CLOVA_OCR_APIGW_INVOKE_URL")
    CLOVA_OCR_API_SECRET = os.getenv("CLOVA_OCR_SECRET_KEY")
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

    # 소스 디렉토리 및 결과 디렉토리 설정
    SOURCE_DIR = Path("source")
    RESULT_DIR = Path("result")

    # 기존 파일 처리
    process_existing_files()

    try:
        # 파일 감시 시작
        event_handler = FileHandler()
        observer = Observer()  # watchdog.observers.polling.PollingObserver() 대신 Observer 사용
        observer.schedule(event_handler, SOURCE_DIR, recursive=False)
        observer.start()

        print(f"파일 감시 시작: {SOURCE_DIR}")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

    # 프로그램 종료 메시지
    print("프로그램을 종료합니다.")


if __name__ == "__main__":
    main()
