import io
import os, string
from google.cloud import vision
import xlsxwriter
import traceback

os.environ["GOOGLE_APPLICATION_CREDENTIALS"]= f"{os.getcwd()}\\auth.json"
clear_console = lambda: os.system('cls')

IMAGE_TITLE = "Image"
SERIAL_TITLE = "Serial"
CODE_TITLE = "Code"

IMAGE_COLUMN = 0
SERIAL_COLUMN = 1
CODE_COLUMN = 2

ALL_FOLDER = '|'

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
FOLDER_NAME = 'test'


def error(message, traceback):
    print(f"""\n
{bcolors.FAIL}Error: {message}{bcolors.ENDC}
{traceback}
    """)
    input('')
    os._exit(1)

def create_work_book(file_name):
    workbook = xlsxwriter.Workbook(f'{file_name}.xlsx')
    return workbook

def create_work_sheet(workbook):
    worksheet = workbook.add_worksheet()
    worksheet.write(0, IMAGE_COLUMN, IMAGE_TITLE)
    worksheet.write(0, SERIAL_COLUMN, SERIAL_TITLE)
    worksheet.write(0, CODE_COLUMN, CODE_TITLE)
    return worksheet

def rename_files(folder_name):
    try:
        files = os.listdir(folder_name)
        for i, file in enumerate(files):
            try:
                filename, file_extension = os.path.splitext(f'{folder_name}\\{file}')
                os.rename(f'{filename}{file_extension}', f'{folder_name}\\__{i+1}{file_extension}')
            except:
                pass
        files = os.listdir(folder_name)
        for i, file in enumerate(files):
            filename, file_extension = os.path.splitext(f'{folder_name}\\{file}')
            os.rename(f'{filename}{file_extension}', f'{folder_name}\\{i+1}{file_extension}')
    except Exception as e:
        tb = traceback.format_exc()
        error(f"Lỗi khi đổi tên file trong thư mục {folder_name}", traceback.format_exc())


def is_serial(text):
    if len(text) > 11:
        if text.startswith(('serial:','095', '083', '1000', '2000', '59')):
            return True
        for c in text:
            if c not in string.digits:
                return False
        return False
    return False

def is_code(text):
    if len(text) > 10:
        if text.startswith(('serial','095', '083', '1000', '2000', '59')):
            return False
        for c in text:
            if c not in string.digits + '-':
                return False
        return True
    return False


def get_texts(file_name):
    try:
        client = vision.ImageAnnotatorClient()
        with io.open(file_name, 'rb') as image_file:
            content = image_file.read()
        image = vision.Image(content=content)
        response = client.text_detection(image=image)
        if response.error.message:
            raise Exception(
                '{}\nFor more info on error messages, check: '
                'https://cloud.google.com/apis/design/errors'.format(
                    response.error.message))
        texts = response.text_annotations
        return texts
    except Exception as e:
        tb = traceback.format_exc()
        error(f"Lỗi khi scan ảnh {file_name}", traceback.format_exc())

def get_data(folder_name):
    data = []
    files = os.listdir(folder_name)
    for i, file in enumerate(files):
        tmp_data = {
            'image': '',
            'serials': [],
            'codes': []
        }
        tmp_data['image'] = file
        texts = get_texts(f'{folder_name}\\{file}')
        for text in texts:
            description = text.description
            if is_serial(description):
                tmp_data['serials'].append(description)
            if is_code(description):
                tmp_data['codes'].append(description)
        data.append(tmp_data)
        print(f'\t{i+1}/{bcolors.OKCYAN}{len(files)}{bcolors.ENDC}', end="\r")
    return data


def choose_folder():
    user_choice = None
    folders = next(os.walk('.'))[1]

    error = ''
    while True:
        clear_console()
        try:
            print(error)
            error = ''
            print("Chọn thư mục:")
            print("0. Tất cả thư mục")
            for i, folder in enumerate(folders):
                print(f"""{i+1}. {folder}""")

            user_choice = int(input("-->  "))
            if user_choice == 0:
                return ALL_FOLDER
            else:
                folder = folders[user_choice-1]
                return folder
        except:
            error = 'Lỗi khi chọn thư mục. Thử lại\n'
            pass
        

def get_excel_file(folder):
    rename_files(folder)

    workbook = create_work_book(folder)
    worksheet = create_work_sheet(workbook)
    data = get_data(folder)

    last_row = 1
    for e in data:
        max_length = max(  len(e['serials']), len(e['codes'])  )
        worksheet.merge_range(last_row, IMAGE_COLUMN, last_row + max_length - 1, IMAGE_COLUMN, e['image'])

        # Write serial
        for i, serial in enumerate(e['serials']):
            i = int(i)
            worksheet.write(last_row+i, SERIAL_COLUMN, serial)
        # Write code
        for i, code in enumerate(e['codes']):
            i = int(i)
            worksheet.write(last_row+i, CODE_COLUMN, code)

        last_row = last_row + max_length + 1
    workbook.close()

def main():
    folder = choose_folder()
    clear_console()
    if folder == ALL_FOLDER:
        print(f"Bắt đầu scan {bcolors.WARNING}tất cả{bcolors.ENDC} thư mục...\n")
        for folder in next(os.walk('.'))[1]:
            print(f"\nĐang scan thư mục: {folder}")
            get_excel_file(folder)
    else:
        print(f"Bắt đầu scan thư mục {bcolors.WARNING}{folder}{bcolors.ENDC}...")
        get_excel_file(folder)

    input(f"\n\n{bcolors.OKBLUE}===== COPYRIGHT NGUYEN QUYNH NHUNG ====={bcolors.ENDC}")
    os._exit(1)
main()


