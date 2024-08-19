import os
import re
import nltk
import random
from openpyxl import Workbook, load_workbook
import subprocess
import os

# Đảm bảo rằng nltk đã được tải về tokenizer
# Thiết lập đường dẫn lưu trữ dữ liệu NLTK
path = os.path.join(os.path.dirname(os.path.abspath(".env")), "nltk")

nltk.data.path.append(path)

# Tải xuống gói 'punkt'
nltk.download("punkt", download_dir=path)

# Biểu thức chính quy để tìm nhiều loại số khác nhau
NUMBER_PATTERN = r"\b\d+(?:,\d+)*(?:\.\d+)?(?:-\d+(?:,\d+)*(?:\.\d+)?)?\b"


def find_year_occurrences(text, file_name):
    occurrences = []
    sentences = nltk.sent_tokenize(text)
    for sentence in sentences:
        matches = re.findall(NUMBER_PATTERN, sentence)
        if matches:
            replaced_sentence = sentence
            count = 0
            for match in matches:
                if count < 2:
                    replaced_sentence = replaced_sentence.replace(match, "bao nhiêu", 1)
                    count += 1
                else:
                    break
            question = f"Câu hỏi: {replaced_sentence.strip('.')}?"
            answer = sentence.strip()
            occurrences.append((question, answer, file_name))
    return occurrences


def process_files_for_questions(folder_path: str, num_questions: int):
    txt_files = []
    print("folder_path", folder_path)
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            print("file", file)
            if file.endswith(".txt"):
                txt_files.append(os.path.join(root, file))
    if not txt_files:
        print("Không tìm thấy file .txt nào trong thư mục.")
        return []
    selected_occurrences = []
    sampled_files = random.sample(txt_files, min(num_questions, len(txt_files)))

    for file_path in sampled_files:
        with open(file_path, "r", encoding="utf-8") as file:
            text = file.read()
            year_occurrences = find_year_occurrences(text, os.path.basename(file_path))
            if year_occurrences:
                selected_occurrences.append(random.choice(year_occurrences))
    remaining_files = list(set(txt_files) - set(sampled_files))

    while len(selected_occurrences) < num_questions and remaining_files:
        file_path = remaining_files.pop()
        with open(file_path, "r", encoding="utf-8") as file:
            text = file.read()
            year_occurrences = find_year_occurrences(text, os.path.basename(file_path))

            if year_occurrences:
                selected_occurrences.append(random.choice(year_occurrences))

    if len(selected_occurrences) < num_questions:
        print(f"Chỉ tìm thấy {len(selected_occurrences)} câu hỏi.")

    return selected_occurrences[:num_questions]


def save_to_excel(occurrences, file_name):
    wb = Workbook()
    ws = wb.active
    ws.append(["Câu hỏi", "Câu trả lời", "Nguồn"])
    for question, answer, source in occurrences:
        ws.append([question, answer, source])
    wb.save(file_name)
    print(f"Đã lưu {len(occurrences)} câu hỏi vào file '{file_name}'.")


if __name__ == "__main__":
    # Ví dụ sử dụng hàm
    folder_path = input("Đường dẫn thư mục file .txt:\t")
    num_questions = int(input("Số lượng câu hỏi:\t"))
    name_file_excel = input("Tên file excel:\t")

    num_questions = 5
    questions = process_files_for_questions(folder_path, num_questions)
    path_file_excel = os.path.join(
        os.path.dirname(os.path.abspath(".env")), f"{name_file_excel}.xlsx"
    )
    save_to_excel(questions, path_file_excel)

    # Lấy đường dẫn của thư mục chứa file Excel
    folder_path = os.path.dirname(path_file_excel)

    # Mở thư mục chứa file
    if os.name == "nt":  # Nếu là Windows
        os.startfile(folder_path)