# python-hwp2image

### 버전 정보  
`Python 3.10.4`  
`pip 22.0.4`

### 실행파일 생성  
`pyinstaller hwp2image.py` - 분할 파일 생성  
`pyinstaller -w -F hwp2image.py` - 단일 파일 생성  
+ 생성된 폴더안에 `FilePathCheckerModule.dll`, `hwpReg.reg` 파일을 넣어줘야 한다.
  + 한글 파일을 열 때 접근 권한을 뜨지 않게 하기 위함.



### 이슈사항  
+ hwp 파일을 수정할 수 있는 프로그램이 깔려있어야 함
+ 필요한 패키지는 import 목록을 보고 직접 설치 해야함
