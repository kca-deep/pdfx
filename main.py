import os
import sys
import json
import time
import base64
import logging
import requests
import mimetypes
import shutil
import traceback
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import fitz  # PyMuPDF
from tqdm import tqdm
import openpyxl
from openpyxl.styles import Alignment

# =============================================================================
# Custom logging handler using tqdm.write
# =============================================================================


class TqdmLoggingHandler(logging.Handler):
    def emit(self, record):
        try:
            msg = self.format(record)
            tqdm.write(msg)
            self.flush()
        except Exception:
            self.handleError(record)


# =============================================================================
# 환경설정 및 디렉토리 생성
# =============================================================================


def setup_environment():
    load_dotenv(override=True)
    config = {
        "SOURCE_DIR": Path(os.getenv("SOURCE_DIR", "source")),
        "MERGED_DIR": Path(os.getenv("MERGED_DIR", "merged")),
        "TEMP_DIR": Path(os.getenv("TEMP_DIR", "temp")),
        "RESULT_DIR": Path(os.getenv("RESULT_DIR", "result")),
        "LOG_DIR": Path(os.getenv("LOG_DIR", "logs")),
        "CLOVA_OCR_API_URL": os.getenv("CLOVA_OCR_APIGW_INVOKE_URL"),
        "CLOVA_OCR_SECRET": os.getenv("CLOVA_OCR_SECRET_KEY"),
        "MAX_FILE_SIZE": int(os.getenv("MAX_FILE_SIZE", 16777216)),
        "OPENAI_API_KEY": os.getenv("OPENAI_API_KEY"),
        "OPENAI_API_URL": "https://api.openai.com/v1/chat/completions",
    }
    return config


def setup_logging():
    config = setup_environment()
    log_dir = config["LOG_DIR"]
    log_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime("%Y%m%d")
    log_filename = log_dir / f"log_{date_str}.log"
    handler = TqdmLoggingHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    file_handler = logging.FileHandler(log_filename, mode="a", encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    )
    logging.basicConfig(level=logging.INFO, handlers=[handler, file_handler])
    logging.info(f"로그 파일 생성됨: {log_filename}")


def ensure_directories(dirs):
    for directory in dirs:
        directory.mkdir(parents=True, exist_ok=True)


# =============================================================================
# 파일 유효성 검사 및 OCR API 호출
# =============================================================================

VALID_EXTENSIONS = [".pdf", ".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp"]


def is_valid_file(file_path, max_size):
    file_path = Path(file_path)
    if not file_path.exists():
        logging.error(f"파일이 존재하지 않습니다: {file_path}")
        return False
    if file_path.suffix.lower() not in VALID_EXTENSIONS:
        logging.error(f"지원되지 않는 파일 형식: {file_path.suffix.lower()}")
        return False
    if file_path.stat().st_size > max_size:
        logging.error(f"파일 크기가 너무 큽니다: {file_path.stat().st_size} 바이트")
        return False
    if file_path.suffix.lower() != ".pdf":
        mime_type, _ = mimetypes.guess_type(str(file_path))
        if not mime_type or not mime_type.startswith("image/"):
            logging.error(f"지원되지 않는 MIME 타입: {mime_type}")
            return False
    return True


def call_clova_ocr_api(file_path, api_url, secret):
    file_path = Path(file_path)
    file_ext = file_path.suffix[1:].lower()
    headers = {"X-OCR-SECRET": secret}
    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
    except Exception:
        logging.exception(f"파일 읽기 실패: {file_path}")
        return None
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
    try:
        response = requests.post(api_url, headers=headers, json=payload)
        if response.status_code == 200:
            logging.info(f"OCR API 호출 성공: {file_path}")
            return response.json()
        else:
            logging.error(
                f"OCR API 호출 실패: {file_path}, 상태코드: {response.status_code}"
            )
            return None
    except Exception:
        logging.exception(f"OCR API 호출 중 오류 발생: {file_path}")
        return None


# =============================================================================
# PDF를 이미지로 변환 (PyMuPDF 사용)
# =============================================================================


def convert_pdf_to_images(pdf_path, temp_dir):
    pdf_path = Path(pdf_path)
    temp_dir = Path(temp_dir)
    process_id = os.getpid()
    current_date = datetime.now().strftime("%Y%m%d")
    temp_subdir = temp_dir / current_date / f"{pdf_path.stem}_{process_id}"
    try:
        if temp_subdir.exists():
            shutil.rmtree(temp_subdir, ignore_errors=True)
        temp_subdir.mkdir(parents=True, exist_ok=True)
        image_paths = []
        with fitz.open(pdf_path) as pdf:
            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                matrix = fitz.Matrix(300 / 72, 300 / 72)
                pix = page.get_pixmap(matrix=matrix)
                image_path = temp_subdir / f"page_{page_num+1}.png"
                pix.save(str(image_path))
                image_paths.append(image_path)
        logging.info(f"PDF 변환 완료: {pdf_path} -> {len(image_paths)} 이미지")
        return image_paths
    except Exception:
        logging.exception(f"PDF 변환 오류: {pdf_path}")
        return []


# =============================================================================
# OCR 결과를 줄 단위 jsonl 객체로 변환 (lineBreak 기준 그룹화, 신뢰도 계산)
# =============================================================================


def ocr_result_to_jsonl_lines(ocr_result):
    lines = []
    images = ocr_result.get("images", [])
    for image in images:
        fields = image.get("fields", [])
        current_line_text = ""
        confidences = []
        line_number = 1
        for field in fields:
            text = field.get("inferText", "").strip()
            conf = field.get("inferConfidence", None)
            if conf is not None:
                confidences.append(conf)
            if text:
                if current_line_text:
                    current_line_text += " " + text
                else:
                    current_line_text = text
            if field.get("lineBreak", False):
                if current_line_text:
                    avg_conf = (
                        round(sum(confidences) / len(confidences), 2)
                        if confidences
                        else None
                    )
                    lines.append(
                        {
                            "line": line_number,
                            "text": current_line_text,
                            "confidence": avg_conf,
                        }
                    )
                    line_number += 1
                    current_line_text = ""
                    confidences = []
        if current_line_text:
            avg_conf = (
                round(sum(confidences) / len(confidences), 2) if confidences else None
            )
            lines.append(
                {"line": line_number, "text": current_line_text, "confidence": avg_conf}
            )
    return lines


# =============================================================================
# 각 파일 단위 OCR 처리 후 jsonl 및 원본 OCR 결과 파일 생성
# (결과는 merged 폴더 내 서브폴더(target_folder)에 저장)
# =============================================================================


def process_file_to_jsonl(file_path, config, target_folder):
    file_path = Path(file_path)
    if not is_valid_file(file_path, config["MAX_FILE_SIZE"]):
        return None

    ocr_results = []
    if file_path.suffix.lower() == ".pdf":
        image_paths = convert_pdf_to_images(file_path, config["TEMP_DIR"])
        if not image_paths:
            logging.error(f"PDF 변환 실패: {file_path}")
            return None
        for img_path in image_paths:
            result = call_clova_ocr_api(
                img_path, config["CLOVA_OCR_API_URL"], config["CLOVA_OCR_SECRET"]
            )
            if result:
                ocr_results.append(result)
            try:
                img_path.unlink()
            except Exception:
                pass
    else:
        result = call_clova_ocr_api(
            file_path, config["CLOVA_OCR_API_URL"], config["CLOVA_OCR_SECRET"]
        )
        if result:
            ocr_results.append(result)

    target_folder.mkdir(parents=True, exist_ok=True)
    raw_ocr_path = target_folder / f"{file_path.stem}.ocr.json"
    try:
        with open(raw_ocr_path, "w", encoding="utf-8") as f:
            json.dump(ocr_results, f, ensure_ascii=False, indent=2)
        logging.info(f"원본 OCR 결과 저장: {raw_ocr_path}")
    except Exception:
        logging.exception(f"원본 OCR 결과 저장 실패: {file_path}")

    all_lines = []
    for ocr_result in ocr_results:
        lines = ocr_result_to_jsonl_lines(ocr_result)
        for obj in lines:
            obj["source_file"] = file_path.name
        all_lines.extend(lines)
    jsonl_path = target_folder / f"{file_path.stem}.jsonl"
    try:
        with open(jsonl_path, "w", encoding="utf-8") as f:
            for obj in all_lines:
                f.write(json.dumps(obj, ensure_ascii=False) + "\n")
        logging.info(f"JSONL 파일 저장: {jsonl_path}")
        return jsonl_path
    except Exception:
        logging.exception(f"JSONL 파일 저장 실패: {file_path}")
        return None


# =============================================================================
# 하위 폴더별로 jsonl 파일 취합 및 순차적 병합 (대상 폴더 기준)
# =============================================================================


def process_subfolder_target(target_folder):
    target_folder = Path(target_folder)
    jsonl_files = list(target_folder.glob("*.jsonl"))
    if not jsonl_files:
        logging.info(f"대상 폴더에 JSONL 파일이 없습니다: {target_folder}")
        return
    jsonl_files.sort(key=lambda f: f.name)
    merged_lines = []
    for jf in jsonl_files:
        try:
            with open(jf, "r", encoding="utf-8") as infile:
                file_lines = [line.strip() for line in infile if line.strip()]
                merged_lines.extend(file_lines)
        except Exception:
            logging.exception(f"JSONL 파일 읽기 실패: {jf}")
    merged_file_path = target_folder / f"{target_folder.name}.jsonl"
    try:
        with open(merged_file_path, "w", encoding="utf-8") as outfile:
            for line in merged_lines:
                outfile.write(line + "\n")
        logging.info(f"병합된 JSONL 파일 생성: {merged_file_path}")
    except Exception:
        logging.exception(f"대상 폴더 병합 실패: {target_folder}")


# =============================================================================
# 하위 폴더 내 모든 파일 처리 및 진행률 표시 (tqdm 활용)
# =============================================================================


def process_all_subfolders(config):
    source_dir = config["SOURCE_DIR"]
    merged_dir = config["MERGED_DIR"]
    if not source_dir.exists():
        logging.error(f"소스 디렉토리가 존재하지 않습니다: {source_dir}")
        return
    for subfolder in source_dir.iterdir():
        if subfolder.is_dir():
            target_folder = merged_dir / subfolder.name
            target_folder.mkdir(parents=True, exist_ok=True)
            files = [
                file
                for file in subfolder.iterdir()
                if file.is_file() and file.suffix.lower() in VALID_EXTENSIONS
            ]
            pbar = tqdm(
                total=len(files),
                desc=f"{subfolder.name} 처리 진행",
                unit="file",
                dynamic_ncols=True,
            )
            for file in files:
                process_file_to_jsonl(file, config, target_folder)
                pbar.update(1)
            pbar.close()
            process_subfolder_target(target_folder)


# =============================================================================
# OpenAI API 응답에서 유효한 JSON 문자열 추출 함수
# =============================================================================


def extract_json_from_response(content):
    start = content.find("{")
    end = content.rfind("}")
    if start != -1 and end != -1 and end > start:
        return content[start : end + 1]
    return content


# =============================================================================
# OpenAI API를 활용해 병합된 JSONL 파일의 내용을 분석하여 지정 항목 추출
# =============================================================================


def analyze_merged_jsonl(merged_text, openai_api_key, openai_api_url):
    prompt = f"""다음은 병합된 jsonl 파일의 내용입니다. 각 줄은 OCR 결과의 일부입니다.
아래 항목들을 추출하여 JSON 형식으로 반환해줘.
항목:
- 접수일자: 문서에서 접수일자를 "YYYY-MM-DD" 형식으로 추출
- 세목: 문서에서 세목 정보를 추출하되, 민간경상보조, 민간위탁사업비, 사업출연금 3개 항목 중 1개로 선택
- 세부사업명: 문서에서 세부사업명을 추출하되, 확인되지 않으면 빈 항목으로 표시
- 내역사업명: 문서에서 사업명 정보를 추출 (내역사업명으로 표시)
- 기관명: 내역사업명 옆에 예금주와 동일한 값으로 표시
- 신뢰도: OCR 결과의 평균 신뢰도를 소수점 두 자리까지 표시
- 연간집행계획액, 기수령액, 월집행계획액, 전월이월액, 당월신청액, 누계: 해당 금액은 jsonl 파일 전체의 총계이며, 숫자만 추출하고 천 단위마다 쉼표(,)를 추가하여 표시
- 은행명: 문서에서 은행명을 추출
- 계좌번호: 문서에서 계좌번호 정보를 추출
- 예금주: 문서에서 예금주 정보를 추출

병합된 jsonl 파일 내용:
{merged_text}
"""
    payload = {
        "model": "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": "당신은 OCR 결과 분석 전문가입니다."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.3,
        "max_tokens": 1000,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {openai_api_key}",
    }
    response = requests.post(openai_api_url, headers=headers, json=payload)
    if response.status_code == 200:
        result = response.json()
        content = result["choices"][0]["message"]["content"]
        try:
            data = json.loads(content)
            return data
        except Exception:
            cleaned = extract_json_from_response(content)
            try:
                data = json.loads(cleaned)
                return data
            except Exception:
                logging.error("JSON 파싱 실패, 응답 내용:")
                logging.error(content)
                return None
    else:
        logging.error(f"OpenAI API 호출 실패, 상태코드: {response.status_code}")
        logging.error(response.text)
        return None


# =============================================================================
# 병합된 JSONL 파일들을 읽어 OpenAI 분석 후 Excel 데이터로 변환 (진행률 포함)
# =============================================================================


def process_merged_jsonl_files(config):
    merged_dir = config["MERGED_DIR"]
    openai_api_key = config["OPENAI_API_KEY"]
    openai_api_url = config["OPENAI_API_URL"]
    data_list = []
    target_files = []
    for subfolder in merged_dir.iterdir():
        if subfolder.is_dir():
            candidate = subfolder / f"{subfolder.name}.jsonl"
            if candidate.exists():
                target_files.append(candidate)
    target_files.sort(key=lambda f: f.name)
    pbar = tqdm(
        total=len(target_files),
        desc="병합 파일 분석 진행",
        unit="file",
        dynamic_ncols=True,
    )
    for mf in target_files:
        try:
            with open(mf, "r", encoding="utf-8") as f:
                merged_text = f.read()
            logging.info(f"분석 시작: {mf.name}")
            structured_data = analyze_merged_jsonl(
                merged_text, openai_api_key, openai_api_url
            )
            if structured_data:
                data_list.append(structured_data)
        except Exception:
            logging.exception(f"병합 파일 처리 실패: {mf}")
        pbar.update(1)
    pbar.close()
    return data_list


# =============================================================================
# Excel 파일로 저장 (RESULT_DIR/output_YYYYMMDD.xlsx, 1행 = 1병합 파일 데이터)
# =============================================================================


def export_excel_from_data(data_list, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "결과"
    # 엑셀 헤더: "세목" 추가, "세목"과 "목세목"은 하나의 '세목'으로 표시,
    # "내역사업명" 옆에 "기관명" 추가 (기관명 = 예금주), 그리고 "신뢰도" 추가
    headers = [
        "접수일자",
        "세목",
        "세부사업명",
        "내역사업명",
        "기관명",
        "신뢰도",
        "연간집행계획액",
        "기수령액",
        "월집행계획액",
        "전월이월액",
        "당월신청액",
        "누계",
        "은행명",
        "계좌번호",
        "예금주",
    ]
    ws.append(headers)
    amount_fields = {
        "연간집행계획액",
        "기수령액",
        "월집행계획액",
        "전월이월액",
        "당월신청액",
        "누계",
    }
    for data in data_list:
        row = []
        for col in headers:
            if col == "기관명":
                value = data.get("예금주", "")
            else:
                value = data.get(col, "")
            if col in amount_fields:
                if isinstance(value, (int, float)):
                    value = f"{int(value):,}"
                else:
                    try:
                        value = f"{int(float(value)):,}"
                    except:
                        value = str(value)
            row.append(value)
        ws.append(row)
    # 금액 열 오른쪽 정렬 처리
    for col_idx, header in enumerate(headers, start=1):
        if header in amount_fields:
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            for cell in ws[col_letter]:
                cell.alignment = Alignment(horizontal="right")
    wb.save(output_path)
    logging.info(f"Excel 파일 저장 완료: {output_path}")


# =============================================================================
# 메인 함수 (파일 감시 제거)
# =============================================================================


def main():
    global config
    config = setup_environment()
    setup_logging()
    ensure_directories(
        [
            config["SOURCE_DIR"],
            config["MERGED_DIR"],
            config["TEMP_DIR"],
            config["RESULT_DIR"],
            config["LOG_DIR"],
        ]
    )

    process_all_subfolders(config)
    logging.info("병합된 JSONL 파일 생성 완료.")

    data_list = process_merged_jsonl_files(config)
    if data_list:
        date_str = datetime.now().strftime("%Y%m%d")
        output_file = config["RESULT_DIR"] / f"output_{date_str}.xlsx"
        export_excel_from_data(data_list, output_file)
    else:
        logging.error("분석된 데이터가 없습니다.")

    logging.info("모든 처리가 완료되었습니다.")


if __name__ == "__main__":
    main()
