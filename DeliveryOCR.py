import os
import re
import sys
from datetime import datetime

import pytesseract
from PIL import Image
from openpyxl import Workbook

def load_normalization_map(filename='normalization_map.txt'):
    norm_map = {}
    if not os.path.exists(filename):
        print(f"{filename} 파일이 없습니다. 수취인 정규화는 생략됩니다.")
        return norm_map
    with open(filename, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if '=' in line:
                k, v = line.split('=', 1)
                norm_map[k.strip()] = v.strip()
    return norm_map

def normalize_recipient(name, norm_map):
    return norm_map.get(name, name)

def extract_date(img):
    width, height = img.size
    crop_area = (0, 0, width//3, height//10)
    cropped = img.crop(crop_area)
    text = pytesseract.image_to_string(cropped, lang='kor+eng')
    date_match = re.search(r'(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})', text)
    if date_match:
        y, m, d = date_match.groups()
        try:
            dt = datetime(int(y), int(m), int(d))
            return dt.strftime('%Y%m%d')
        except:
            pass
    return None

def ocr_with_boxes(img):
    data = pytesseract.image_to_data(img, lang='kor+eng', output_type=pytesseract.Output.DICT)
    n_boxes = len(data['level'])
    results = []
    for i in range(n_boxes):
        text = data['text'][i].strip()
        if text:
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            results.append({'text': text, 'left': x, 'top': y, 'width': w, 'height': h})
    return results

def is_registration_number(text):
    return re.match(r'^\d{4}[-]?\d{4}[-]?\d{4}$', text) is not None

def sort_order(records, img_width, img_height):
    mid_x = img_width // 2
    mid_y = img_height // 2

    def region_order(r):
        x = r['left']
        y = r['top']
        if x < mid_x and y < mid_y:
            return 0
        elif x < mid_x and y >= mid_y:
            return 1
        elif x >= mid_x and y < mid_y:
            return 2
        else:
            return 3

    return sorted(records, key=lambda r: (region_order(r), r['top'], r['left']))

def process_image(image_path, norm_map):
    img = Image.open(image_path)
    date_str = extract_date(img)
    if not date_str:
        date_str = datetime.now().strftime('%Y%m%d')

    ocr_results = ocr_with_boxes(img)
    img_width, img_height = img.size

    records = []

    for i, r in enumerate(ocr_results):
        text = r['text']
        if is_registration_number(text.replace(' ', '').replace('_', '')):
            reg_num = text.replace(' ', '').replace('_', '')
            recip = None
            for j in range(i+1, min(i+6, len(ocr_results))):
                nr = ocr_results[j]
                if abs(nr['top'] - r['top']) < 20 and nr['left'] > r['left']:
                    recip = nr['text']
                    break
                if nr['top'] > r['top'] and abs(nr['left'] - r['left']) < 50:
                    recip = nr['text']
                    break
            if recip:
                recip_norm = normalize_recipient(recip, norm_map)
                records.append({'reg_num': reg_num, 'recipient': recip_norm, 'left': r['left'], 'top': r['top']})

    if not records:
        print("등기번호 및 수취인 정보를 찾지 못했습니다.")
        return None, None

    sorted_records = sort_order(records, img_width, img_height)
    for idx, rec in enumerate(sorted_records, start=1):
        rec['순번'] = idx

    return date_str, sorted_records

def save_to_excel(date_str, records, output_dir='output'):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    filename = f"{date_str}_우체국 다량우편물 배달증.xlsx"
    filepath = os.path.join(output_dir, filename)

    wb = Workbook()
    ws = wb.active
    ws.title = '배달증 결과'

    ws.append(['순번', '등기번호', '수취인'])
    for rec in records:
        ws.append([rec['순번'], rec['reg_num'], rec['recipient']])

    wb.save(filepath)
    print(f"결과가 저장되었습니다: {filepath}")
    return filepath

def main():
    if len(sys.argv) < 2:
        print("사용법: DeliveryOCR.exe <이미지 파일 경로>")
        return

    image_path = sys.argv[1]
    if not os.path.exists(image_path):
        print("이미지 파일을 찾을 수 없습니다.")
        return

    norm_map = load_normalization_map()
    date_str, records = process_image(image_path, norm_map)
    if records:
        save_to_excel(date_str, records)

if __name__ == '__main__':
    main()
