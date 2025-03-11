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
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv
import openpyxl
import fitz  # PyMuPDF
from PIL import Image
import io
import traceback
from tqdm import tqdm

# =============================================================================
# 초기 설정 및 디렉토리 생성
# =============================================================================


def setup_environment():
    load_dotenv(override=True)  # 기존 환경 변수 덮어쓰기
    config = {
        "LOG_LEVEL": os.getenv("LOG_LEVEL", "INFO"),
        "MAX_FILE_SIZE": int(os.getenv("MAX_FILE_SIZE", 16777216)),
        "TEMP_DIR": Path(os.getenv("TEMP_DIR", "temp")),
        "SOURCE_DIR": Path(os.getenv("SOURCE_DIR", "source")),
        "RESULT_DIR": Path(os.getenv("RESULT_DIR", "result")),
        "CLOVA_OCR_API_URL": os.getenv("CLOVA_OCR_APIGW_INVOKE_URL"),
        "CLOVA_OCR_SECRET": os.getenv("CLOVA_OCR_SECRET_KEY"),
        "OPENAI_API_KEY": os.getenv("OPENAI_API_KEY"),
        "OPENAI_API_URL": "https://api.openai.com/v1/chat/completions",
    }
    return config


def setup_logging():
    log_dir = Path("logs")
    log_dir.mkdir(parents=True, exist_ok=True)
    current_date = datetime.now().strftime("%Y%m%d")
    log_filename = log_dir / f"pdfx_{current_date}.log"
    logging.basicConfig(
        level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.FileHandler(str(log_filename), mode="a", encoding="utf-8")],
    )
    logging.info(f"로그 파일 경로: {log_filename}")
    return log_filename


def ensure_directories(dirs):
    for directory in dirs:
        directory.mkdir(exist_ok=True, parents=True)


# =============================================================================
# API 사용량 추적 클래스
# =============================================================================


class APIUsageTracker:
    def __init__(self, exchange_rate=1450):
        self.total_tokens = 0
        self.total_prompt_tokens = 0
        self.total_completion_tokens = 0
        self.total_cost_usd = 0.0
        self.exchange_rate = exchange_rate  # 1 USD = 1450 KRW
        self.model_costs = {
            "gpt-4o-mini": {"input": 0.00000015, "output": 0.0000006},
            "gpt-4o": {"input": 0.000003, "output": 0.000006},
            "gpt-3.5-turbo": {"input": 0.0000005, "output": 0.0000015},
        }
        self.default_model = "gpt-4o-mini"

    def count_tokens(self, text):
        words = text.split()
        word_count = len(words)
        char_count = len(text)
        has_korean = any(0xAC00 <= ord(char) <= 0xD7A3 for char in text)
        if has_korean:
            return int(char_count * 1.5)
        else:
            return int(word_count * 0.75)

    def update_usage(self, prompt_tokens, completion_tokens, model=None):
        model = model if model in self.model_costs else self.default_model
        self.total_prompt_tokens += prompt_tokens
        self.total_completion_tokens += completion_tokens
        self.total_tokens += prompt_tokens + completion_tokens
        prompt_cost = prompt_tokens * self.model_costs[model]["input"]
        completion_cost = completion_tokens * self.model_costs[model]["output"]
        cost_usd = prompt_cost + completion_cost
        self.total_cost_usd += cost_usd
        cost_krw = cost_usd * self.exchange_rate
        total_cost_krw = self.total_cost_usd * self.exchange_rate
        logging.info(
            f"API 사용량 업데이트: 모델={model}, 입력 토큰={prompt_tokens}, 출력 토큰={completion_tokens} | "
            f"토큰 비용 (1000토큰당): 입력=${self.model_costs[model]['input']*1000:.4f}, 출력=${self.model_costs[model]['output']*1000:.4f} | "
            f"현재 호출 비용: ${cost_usd:.6f} (₩{cost_krw:.2f}) | 누적 토큰={self.total_tokens} / 누적 비용: ${self.total_cost_usd:.6f} (₩{total_cost_krw:.2f})"
        )
        return {
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": prompt_tokens + completion_tokens,
            "cost_usd": cost_usd,
            "cost_krw": cost_krw,
        }


# =============================================================================
# OpenAI API를 활용한 OCR 결과 분석
# =============================================================================


def analyze_ocr_with_openai(ocr_result, api_tracker, OPENAI_API_KEY, OPENAI_API_URL):
    if not OPENAI_API_KEY:
        logging.error("OpenAI API 키가 설정되지 않았습니다.")
        return None

    # OCR 결과 텍스트 추출
    all_text = ""
    images = ocr_result.get("images", [])
    if images:
        for field in images[0].get("fields", []):
            if "inferText" in field:
                y_coord = (
                    field.get("boundingPoly", {}).get("vertices", [{}])[0].get("y", 0)
                )
                all_text += f"행 {int(y_coord/10)}: {field.get('inferText','')}\n"

    if not all_text:
        logging.warning("분석할 텍스트가 없습니다.")
        return None

    prompt = f"""
다음은 OCR로 추출한 텍스트입니다. 이 텍스트에서 다음 정보를 추출해주세요:

1. 신청일자 정보: 문서에서 신청일자, 작성일자, 기안일자 등을 찾아 'YYYY-MM-DD' 형식으로 변환
2. 5~7행에 있는 금액 데이터: 연간집행, 기수령액, 계획액, 월집행, 이월액, 전월, 신청액, 당월, 누계
3. 22~24행에 있는 은행 정보: 은행명, 계좌번호, 예금주
4. 첫번째 행에서 '기금구분': 방송통신발전기금, 정보통신진흥기금 중 문서에 있는 값
5. '목/세목': 민간위탁사업비, 민간경상보조, 사업출연금 중 문서에 있는 값

추출한 정보를 다음 JSON 형식으로 반환해주세요:
{{
    "금액_데이터": [
        {{"항목": "연간집행계획액", "금액": "금액"}},
        {{"항목": "기수령액", "금액": "금액"}},
        {{"항목": "월집행계획액", "금액": "금액"}},
        {{"항목": "전월이월액", "금액": "금액"}},
        {{"항목": "당월신청액", "금액": "금액"}},
        {{"항목": "누계", "금액": "금액"}}
    ],
    "은행_정보": [
        {{"항목": "은행명", "값": "은행명"}},
        {{"항목": "계좌번호", "값": "계좌번호"}},
        {{"항목": "예금주", "값": "예금주"}}
    ],
    "신청일자": "YYYY-MM-DD",
    "기금구분": "기금구분",
    "목세목": "목세목"
}}
OCR 텍스트:
{all_text}
"""
    model = "gpt-4o-mini"
    prompt_tokens = api_tracker.count_tokens(prompt)
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "당신은 OCR 결과에서 정보를 추출하는 전문가입니다.",
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.3,
        "max_tokens": 1000,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {OPENAI_API_KEY}",
    }

    try:
        response = requests.post(
            OPENAI_API_URL, headers=headers, data=json.dumps(payload)
        )
        if response.status_code == 200:
            result = response.json()
            usage = result.get("usage", {})
            prompt_tokens_used = usage.get("prompt_tokens", prompt_tokens)
            completion_tokens = usage.get("completion_tokens", 0)
            api_tracker.update_usage(prompt_tokens_used, completion_tokens, model)
            response_text = result["choices"][0]["message"]["content"]
            json_start = response_text.find("{")
            json_end = response_text.rfind("}") + 1
            if json_start >= 0 and json_end > json_start:
                return json.loads(response_text[json_start:json_end])
            else:
                logging.error("응답에서 JSON 형식을 찾을 수 없습니다.")
                logging.error(f"원본 응답: {response_text}")
                return None
        else:
            logging.error(
                f"OpenAI API 요청 실패: {response.status_code} - {response.text}"
            )
            return None
    except Exception:
        logging.exception("OpenAI API 호출 중 오류 발생")
        return None


# =============================================================================
# Excel 파일 저장 함수 (헤더 및 값 추가) - 헤더 중복 방지 개선
# =============================================================================


def generate_header(analyzed_data):
    header = []
    # 파일 정보 헤더
    for item in analyzed_data.get("파일_정보", []):
        header.append(item.get("항목", ""))
    header.append("")  # 구분선
    # 금액 데이터 헤더
    for item in analyzed_data.get("금액_데이터", []):
        header.append(item.get("항목", ""))
    header.append("")  # 구분선
    # 은행 정보 헤더
    for item in analyzed_data.get("은행_정보", []):
        header.append(item.get("항목", ""))
    return header


def generate_row(analyzed_data):
    row = []
    # 파일 정보 값
    for item in analyzed_data.get("파일_정보", []):
        row.append(item.get("값", ""))
    row.append("")
    # 금액 데이터 값
    for item in analyzed_data.get("금액_데이터", []):
        value = item.get("금액", "")
        cleaned = value.replace(",", "").replace(" ", "")
        row.append(int(cleaned) if cleaned.isdigit() else value)
    row.append("")
    # 은행 정보 값
    for item in analyzed_data.get("은행_정보", []):
        row.append(item.get("값", ""))
    return row


def export_to_structured_excel(analyzed_data, output_path):
    if not analyzed_data:
        logging.error("내보낼 분석 데이터가 없습니다.")
        return False

    try:
        if output_path.exists():
            wb = openpyxl.load_workbook(output_path)
            ws = (
                wb["추출 데이터"]
                if "추출 데이터" in wb.sheetnames
                else wb.create_sheet("추출 데이터")
            )
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "추출 데이터"

        # 헤더 행이 비어있다면 생성
        first_row = [cell.value for cell in ws[1]]
        if not any(first_row):
            header = generate_header(analyzed_data)
            for col, header_text in enumerate(header, start=1):
                cell = ws.cell(row=1, column=col)
                cell.value = header_text
                cell.font = openpyxl.styles.Font(bold=True)
                cell.fill = openpyxl.styles.PatternFill(
                    start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"
                )

        # 데이터 행 추가
        row = generate_row(analyzed_data)
        ws.append(row)

        # 열 너비 자동 조정
        for column in ws.columns:
            max_length = max(
                (len(str(cell.value)) for cell in column if cell.value), default=0
            )
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(column[0].column)
            ].width = (max_length + 2) * 1.2

        wb.save(output_path)
        logging.info(f"Excel 파일 저장 완료: {output_path}")
        return True
    except Exception:
        logging.exception("Excel 파일 생성 중 오류 발생")
        return False


# =============================================================================
# 파일 및 PDF 처리 관련 함수
# =============================================================================

ALL_SUPPORTED_EXTENSIONS = [".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".pdf"]


def is_valid_file(file_path, MAX_FILE_SIZE):
    file_path = Path(file_path)
    if not file_path.exists():
        logging.error(f"파일이 존재하지 않습니다: {file_path}")
        return False
    if file_path.suffix.lower() not in ALL_SUPPORTED_EXTENSIONS:
        logging.error(f"지원되지 않는 파일 형식: {file_path.suffix.lower()}")
        return False
    if file_path.stat().st_size > MAX_FILE_SIZE:
        logging.error(f"파일 크기가 너무 큽니다: {file_path.stat().st_size} 바이트")
        return False
    if file_path.suffix.lower() != ".pdf":
        mime_type, _ = mimetypes.guess_type(str(file_path))
        if not mime_type or not mime_type.startswith("image/"):
            logging.error(f"지원되지 않는 MIME 타입: {mime_type}")
            return False
    return True


def convert_pdf_to_images(pdf_path, temp_dir, pbar=None):
    pdf_path = Path(pdf_path)
    logging.info(f"PDF 변환 시작: {pdf_path}")
    process_id = os.getpid()
    current_date = datetime.now().strftime("%Y%m%d")
    temp_subdir = temp_dir / current_date / f"{pdf_path.stem}_{process_id}"
    try:
        if temp_subdir.exists():
            import shutil

            shutil.rmtree(temp_subdir, ignore_errors=True)
        temp_subdir.mkdir(parents=True, exist_ok=True)
        image_paths = []
        with fitz.open(pdf_path) as pdf_document:
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))
                image_path = temp_subdir / f"page_{page_num + 1}.png"
                if image_path.exists():
                    image_path.unlink(missing_ok=True)
                try:
                    pix.save(str(image_path))
                    image_paths.append(image_path)
                    if pbar:
                        pbar.update(1)
                        pbar.refresh()
                except Exception:
                    logging.exception(f"이미지 저장 실패 (페이지 {page_num+1})")
        if not image_paths:
            logging.error("PDF 변환 실패: 이미지가 생성되지 않음")
        else:
            logging.info(f"PDF 변환 완료: {len(image_paths)}개 이미지 생성")
        return image_paths
    except Exception:
        logging.exception("PDF 변환 중 오류 발생")
        return []


def call_clova_ocr_api(file_path, CLOVA_OCR_API_URL, CLOVA_OCR_SECRET):
    file_path = Path(file_path)
    file_ext = file_path.suffix[1:].lower()
    headers = {"X-OCR-SECRET": CLOVA_OCR_SECRET}
    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
    except Exception:
        logging.exception(f"파일 읽기 실패: {file_path}")
        return None

    is_custom_api = "custom" in CLOVA_OCR_API_URL
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            if is_custom_api:
                timestamp = int(time.time() * 1000)
                message = {
                    "version": "V2",
                    "requestId": f"request_{timestamp}",
                    "timestamp": timestamp,
                    "images": [{"format": file_ext, "name": file_path.name}],
                }
                files = {
                    "file": (file_path.name, file_data, f"image/{file_ext}"),
                    "message": (None, json.dumps(message)),
                }
                response = requests.post(
                    CLOVA_OCR_API_URL, headers=headers, files=files
                )
                if response.status_code == 200:
                    logging.info(f"OCR API 호출 성공 (custom): {file_path}")
                    return response.json()
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
                json_headers = {**headers, "Content-Type": "application/json"}
                response = requests.post(
                    CLOVA_OCR_API_URL, headers=json_headers, json=json_data
                )
                if response.status_code == 200:
                    logging.info(f"JSON 형식 OCR API 호출 성공 (custom): {file_path}")
                    return response.json()
            else:
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
                headers["Content-Type"] = "application/json"
                response = requests.post(
                    CLOVA_OCR_API_URL, headers=headers, json=payload
                )
                if response.status_code == 200:
                    logging.info(f"OCR API 호출 성공: {file_path}")
                    return response.json()
            retry_count += 1
            if retry_count < max_retries:
                wait_time = 2**retry_count
                logging.info(f"{wait_time}초 후 재시도: {file_path}")
                time.sleep(wait_time)
        except requests.exceptions.RequestException:
            retry_count += 1
            if retry_count < max_retries:
                wait_time = 2**retry_count
                logging.info(
                    f"네트워크 오류 발생. {wait_time}초 후 재시도: {file_path}"
                )
                time.sleep(wait_time)
        except Exception:
            logging.exception(f"OCR API 호출 중 예외 발생: {file_path}")
            return None
    logging.error(f"OCR API 호출 실패 (최대 재시도 초과): {file_path}")
    return None


def process_ocr_result(ocr_result, file_path, result_dir, pbar=None):
    file_base = file_path.stem
    json_filename = f"{file_base}_{os.getpid()}.json"
    json_path = result_dir / json_filename
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(ocr_result, f, ensure_ascii=False, indent=2)
        logging.info(f"JSON 결과 저장: {json_path}")
    except Exception:
        logging.exception("JSON 결과 저장 실패")
        json_path = Path.cwd() / json_filename
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(ocr_result, f, ensure_ascii=False, indent=2)
    if pbar:
        pbar.set_description(f"데이터 분석 중: {file_path.name}")
        pbar.refresh()
    analyzed_data = analyze_ocr_with_openai(
        ocr_result, api_tracker, config["OPENAI_API_KEY"], config["OPENAI_API_URL"]
    )
    return analyzed_data


def add_file_info(analyzed_data, file_path, current_datetime, page_info=None):
    if not analyzed_data:
        return analyzed_data
    analyzed_data.setdefault("파일_정보", [])
    if analyzed_data.get("신청일자"):
        analyzed_data["파일_정보"].append(
            {"항목": "신청일자", "값": analyzed_data["신청일자"]}
        )
    if analyzed_data.get("기금구분"):
        analyzed_data["파일_정보"].append(
            {"항목": "기금구분", "값": analyzed_data["기금구분"]}
        )
    if analyzed_data.get("세부사업"):
        analyzed_data["파일_정보"].append(
            {"항목": "세부사업", "값": analyzed_data["세부사업"]}
        )
    if analyzed_data.get("목세목"):
        analyzed_data["파일_정보"].append(
            {"항목": "목세목", "값": analyzed_data["목세목"]}
        )
    analyzed_data["파일_정보"].append({"항목": "원본파일명", "값": file_path.name})
    analyzed_data["파일_정보"].append({"항목": "처리시간", "값": current_datetime})
    if page_info:
        analyzed_data["파일_정보"].append({"항목": "페이지번호", "값": page_info})
    return analyzed_data


def save_to_excel(
    analyzed_data, common_output_path, result_dir, file_base, current_date
):
    if not analyzed_data:
        return False
    if export_to_structured_excel(analyzed_data, common_output_path):
        logging.info(f"공통 Excel 파일에 저장됨: {common_output_path}")
        return True
    else:
        try:
            alt_path = Path.cwd() / "output.xlsx"
            if export_to_structured_excel(analyzed_data, alt_path):
                logging.info(f"대체 경로 저장 성공: {alt_path}")
                return True
        except Exception:
            logging.exception("대체 경로 Excel 저장 실패")
        try:
            individual_path = result_dir / f"output_{file_base}_{current_date}.xlsx"
            if export_to_structured_excel(analyzed_data, individual_path):
                logging.info(f"개별 Excel 파일 저장 성공: {individual_path}")
                return True
        except Exception:
            logging.exception("개별 Excel 저장 실패")
    return False


def process_file(file_path, config, pbar=None):
    file_path = Path(file_path)
    if not is_valid_file(file_path, config["MAX_FILE_SIZE"]):
        logging.error(f"파일 유효성 검사 실패: {file_path}")
        return

    try:
        file_base = file_path.stem
        current_date = datetime.now().strftime("%Y%m%d")
        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")
        result_dir = config["RESULT_DIR"] / current_date
        result_dir.mkdir(parents=True, exist_ok=True)
        common_output_path = result_dir / "output.xlsx"

        if file_path.suffix.lower() == ".pdf":
            if pbar:
                pbar.set_description(f"PDF 변환 중: {file_path.name}")
                pbar.refresh()
            image_paths = convert_pdf_to_images(file_path, config["TEMP_DIR"], pbar)
            if not image_paths:
                logging.error(f"PDF 변환 실패: {file_path}")
                return
            for i, image_path in enumerate(image_paths):
                ocr_result = call_clova_ocr_api(
                    image_path, config["CLOVA_OCR_API_URL"], config["CLOVA_OCR_SECRET"]
                )
                if pbar:
                    pbar.update(1)
                    pbar.refresh()
                if not ocr_result:
                    logging.error(f"OCR API 호출 실패: {image_path}")
                    continue
                analyzed_data = process_ocr_result(
                    ocr_result, file_path, result_dir, pbar
                )
                if analyzed_data:
                    page_info = f"{i+1}/{len(image_paths)}"
                    analyzed_data = add_file_info(
                        analyzed_data, file_path, current_datetime, page_info
                    )
                    save_to_excel(
                        analyzed_data,
                        common_output_path,
                        result_dir,
                        file_base,
                        current_date,
                    )
        else:
            logging.info(f"이미지 파일 처리 시작: {file_path}")
            if pbar:
                pbar.set_description(f"OCR 처리 중: {file_path.name}")
                pbar.refresh()
            ocr_result = call_clova_ocr_api(
                file_path, config["CLOVA_OCR_API_URL"], config["CLOVA_OCR_SECRET"]
            )
            if pbar:
                pbar.update(1)
                pbar.refresh()
            if not ocr_result:
                logging.error(f"OCR API 호출 실패: {file_path}")
                return
            analyzed_data = process_ocr_result(ocr_result, file_path, result_dir, pbar)
            if analyzed_data:
                analyzed_data = add_file_info(
                    analyzed_data, file_path, current_datetime
                )
                save_to_excel(
                    analyzed_data,
                    common_output_path,
                    result_dir,
                    file_base,
                    current_date,
                )
        logging.info(f"파일 처리 완료: {file_path}")
    except Exception:
        logging.exception(f"파일 처리 중 오류 발생: {file_path}")


def process_existing_files(config):
    source_dir = config["SOURCE_DIR"]
    source_dir.mkdir(parents=True, exist_ok=True)
    # 하위 폴더까지 재귀적으로 검색 (rglob 사용)
    files = [
        f for f in source_dir.rglob("*") if f.suffix.lower() in ALL_SUPPORTED_EXTENSIONS
    ]
    if not files:
        return
    logging.info(f"총 {len(files)}개 파일 발견")
    total_steps = 0
    for file_path in files:
        if file_path.suffix.lower() == ".pdf":
            try:
                with fitz.open(file_path) as pdf_document:
                    total_steps += len(pdf_document) * 2
            except Exception:
                logging.exception("PDF 페이지 수 확인 실패")
                total_steps += 2
        else:
            total_steps += 1
    with tqdm(total=total_steps, desc="파일 처리", ncols=80, file=sys.stdout) as pbar:
        for file_path in files:
            if file_path.is_file():
                try:
                    process_file(file_path, config, pbar)
                except Exception:
                    logging.exception(f"파일 처리 중 오류 발생: {file_path}")


# =============================================================================
# 파일 감시 및 메인 함수
# =============================================================================


class FileHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        file_path = Path(event.src_path)
        try:
            process_file(file_path, config)
        except Exception:
            logging.exception(f"파일 처리 중 오류 발생: {file_path}")


def main():
    global config, api_tracker
    config = setup_environment()
    log_file = setup_logging()
    ensure_directories([config["TEMP_DIR"], config["SOURCE_DIR"], config["RESULT_DIR"]])

    print(f"PDF-X 프로그램 시작 (로그 파일: {log_file})")
    print(f"소스 디렉토리: {config['SOURCE_DIR'].absolute()}")
    print(f"결과 디렉토리: {config['RESULT_DIR'].absolute()}")

    # 필수 환경 변수 확인
    for var in ["CLOVA_OCR_API_URL", "CLOVA_OCR_SECRET", "OPENAI_API_KEY"]:
        if not config.get(var):
            logging.error(f"필수 환경 변수 {var}가 설정되지 않았습니다.")
            sys.exit(1)

    global api_tracker
    api_tracker = APIUsageTracker(exchange_rate=1450)
    logging.info(f"CLOVA OCR API URL: {config['CLOVA_OCR_API_URL']}")
    if config["CLOVA_OCR_SECRET"]:
        secret = config["CLOVA_OCR_SECRET"]
        logging.info(
            f"CLOVA OCR SECRET: {secret[:4]}...{secret[-4:]} (길이: {len(secret)})"
        )

    process_existing_files(config)

    try:
        event_handler = FileHandler()
        # recursive=True로 설정하여 소스 폴더의 하위 폴더도 감시
        observer = Observer()
        observer.schedule(event_handler, str(config["SOURCE_DIR"]), recursive=True)
        observer.start()
        print(f"파일 감시 시작: {config['SOURCE_DIR']}")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        observer.join()
        print("프로그램을 종료합니다.")


if __name__ == "__main__":
    main()
