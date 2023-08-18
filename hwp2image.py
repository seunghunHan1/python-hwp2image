import os
import win32com.client as win32
import win32api
from tkinter.filedialog import askopenfilename, askdirectory
import tkinter
import cv2
import glob
import clipboard
import shutil


class Hwp:
    def __init__(self):
        self.current_directory = os.getcwd()

    def create_result_folder(self):
        try:
            os.mkdir(os.path.join(self.current_directory, 'result'))
        except FileExistsError:
            print("이미 result 폴더가 있습니다.")

    def create_cropped_image(self, file_path: str):
        hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
        hwp.RegisterModule('FilePathCheckDLL',
                           'FilePathCheckerModule')  # 한글 보안모듈 설정
        hwp.XHwpWindows.Item(0).Visible = True
        hwp.Open(file_path)
        hwp.MovePos(0)

        if (hwp.PageCount % 3 != 0):
            hwp.Quit()
            win32api.MessageBox(
                0, file_path + " 파일의 총 페이지 수가 3의 배수가 아닙니다", "총 페이지 수가 3의 배수가 아닙니다", 16)
            return

        # question_code로 폴더 생성 작업
        file_codes = []
        for i in range(hwp.PageCount):
            if (i + 1) % 3 == 0:
                hwp.Run('CopyPage')
                question_code = clipboard.paste().strip()
                file_codes.append(question_code)
                try:
                    os.mkdir(os.path.join(
                        self.current_directory, 'result', question_code))
                except FileExistsError:
                    shutil.rmtree(os.path.join(
                        self.current_directory, 'result', question_code))
                    os.mkdir(os.path.join(
                        self.current_directory, 'result', question_code))
                    print('이미' + question_code + '폴더가 있습니다.')
            hwp.Run('MovePageDown')

        # 문서 최상단으로 이동
        hwp.MovePos(0)

        for i in range(hwp.PageCount):
            # 3의 배수 페이지일 경우 문제나 해설이 아니기 때문에 pass
            if (i + 1) % 3 == 0:
                hwp.Run('MovePageDown')
                continue

            # 이미지의 파일 이름 선언
            file_name = file_codes[int(i / 3)] + \
                ('_description' if ((i + 1) - (3 * int(i / 3))) %
                 2 == 0 else '_question')

            image_path = os.path.join(self.current_directory,
                                      'result', file_codes[int(i / 3)])

            hwp.CreatePageImage(os.path.join(
                image_path, file_name), i, 'bmp')

            os.rename(os.path.join(
                image_path, file_name + '.bmp'), os.path.join(
                image_path, file_name + '.png'))

            # 다음 페이지로 이동
            hwp.Run('MovePageDown')

        hwp.Quit()

        # 만들어지 이미지에 사용되지 않는 흰여백을 제거 하고 덮어씌운다
        for question_code in file_codes:
            image_path = os.path.join(self.current_directory,
                                      'result', question_code)
            Cv2().cropped_image_action(image_path)

    def target_folder_create_cropped_image(self, folder_path: str):
        hwp_files = glob.glob(os.path.join(
            folder_path, '**/*.hwp'), recursive=True)

        for file_path in hwp_files:
            self.create_cropped_image(file_path)


class Cv2:
    def cropped_image_action(self, question_path: str):
        os.chdir(question_path)
        fileEx = r'.png'
        png_list = [file for file in os.listdir() if file.endswith(fileEx)]

        for file_name in png_list:
            padding = 10  # 이미지 주위 흰 여백 값
            image = cv2.imread(file_name)
            gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            edges = cv2.Canny(
                gray_image, threshold1=100, threshold2=200)
            contours, _ = cv2.findContours(
                edges.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            x_min, y_min = image.shape[1], image.shape[0]
            x_max, y_max = 0, 0

            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                x_min = min(x_min, x)
                y_min = min(y_min, y)
                x_max = max(x_max, x + w)
                y_max = max(y_max, y + h)
            x_min = max(0, x_min - padding)
            y_min = max(0, y_min - padding)
            x_max = min(image.shape[1], x_max + padding)
            y_max = min(image.shape[0], y_max + padding)
            cropped_image = image[y_min:y_max, x_min:x_max]

            cv2.imwrite(file_name, cropped_image)


class Tinker:
    def __init__(self):
        self.window = tkinter.Tk()

        self.create_tinker()

    def create_tinker(self):
        self.window.title("HWP 문제 생성")
        self.window.geometry("400x200")
        self.window.resizable(None, None)

        file_select_button = tkinter.Button(self.window, overrelief="solid", width=30,
                                            command=self.file_select_action, repeatdelay=1000, repeatinterval=100, text='파일 선택 - 선택한 hwp만 작업')
        file_select_button.pack(side='top', ipady=3, pady=10)

        folder_select_button = tkinter.Button(self.window, overrelief="solid", width=35,
                                              command=self.folder_select_action, repeatdelay=1000, repeatinterval=100, text='폴더 선택 - 선택 폴더 아래 모든 hwp 작업')
        folder_select_button.pack(side='top', ipady=5)

        label1 = tkinter.Label(
            self.window, text="참고사항")
        label1.pack(side='top', pady=10)

        label2 = tkinter.Label(
            self.window, text="결과물은 해당 실행 파일과 같은 경로의 result 폴더 아래에 있습니다.")
        label2.pack(side='top')

        label3 = tkinter.Label(
            self.window, text="이미 같은 결과물이 있을 경우에는 파일이 덮어 씌워집니다.")
        label3.pack(side='top', pady=5)

        self.window.mainloop()

    def show_clear_msg_box(self):
        win32api.MessageBox(0, "hwp 파일 문제 작업 완료", "작업 완료", 1)

    def file_select_action(self):
        win32api.MessageBox(0, "열려 있는 hwp 파일이 있다면 닫아주세요.", "실행 전", 32)
        hwp = Hwp()
        file_path = askopenfilename(title="hwp 파일을 선택하세요.",
                                    initialdir=os.getcwd(),
                                    filetypes=[("*.hwp", "*.hwp *.hwpx")])

        if file_path:
            hwp.create_result_folder()
            hwp.create_cropped_image(file_path)
            self.show_clear_msg_box()

    def folder_select_action(self):
        win32api.MessageBox(0, "열려 있는 hwp 파일이 있다면 닫아주세요.", "실행 전", 32)
        dir_path = askdirectory(title='hwp 파일이 있는 폴더를 선택하세요.')
        if dir_path:
            hwp = Hwp()
            hwp.create_result_folder()
            hwp.target_folder_create_cropped_image(dir_path)
            self.show_clear_msg_box()


def main():
    Tinker()


if __name__ == '__main__':
    main()
