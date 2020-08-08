import os
import subprocess
from os import chdir
from os.path import exists

# 실행 예시:
# python txt_to_hwp.py  --> 받아올 텍스트파일입력(txt)  --> 결과물 저장할파일입력(txt)

# 이 파이썬 코드는 텍스트 형식의 내용을 받아 속기사가 작성하는 회의록 형식의 한글파일로 만들어주는 코드입니다.
# 기존의 한글변환 자바코드가 있기때문에, 그 코드를 명령어 방식으로 파이썬안에서 실행시켰습니다.(명령어를 이용해 자바의 jar파일을 실행시키는 방식)
# 명령어를 실행시키기 위해 os와 subprocess라는 라이브러리를 사용했습니다.
# 주석은 해당하는 코드 밑에 작성했습니다.
# 순서
# 1: 초기 디렉토리 설정
# 2: jar파일을 실행하는 함수 정의(execute_jar)
#   2.1: 텍스트파일과 결과파일 입력 및 예외처리
#   2.2: 명령어를 통한 jar 파일 실행(중요)
# 3: 실행


# 1)
modified_osgetcwd = os.getcwd().replace("\\", "/")
# 현재 디렉토리를 사용하기 위해 변수에 저장

java_src_dir = "test_hwp"
jar_src_dir = "test_hwp/out/artifacts/test_jar/test.jar"
# jar파일을 실행시키기 위한 디렉토리, 명령어로 java 실행을 위해서는 디렉토리가 필요

def move_dir():
    if exists(java_src_dir):
        chdir(java_src_dir)
        # 디렉토리 이동을 위한 함수

# 2)
def execute_jar(java_file):
    # 2.1)
    print("Enter the source file name")
    source_name = input()

    try:
        f = open(source_name, 'rt', encoding='utf-8')
        # input으로 받은 텍스트가 실제로 존재하고 열리는 파일인지 확인
    except FileNotFoundError:
        print("Cannot find the file")
        return

    print("Enter the name of the result file")
    result_name = input()

    # 2.2)
    subprocess.check_call(['java', "-jar", "-Dfile.encoding=UTF-8", java_file, source_name, result_name])
    # 명령어로 jar파일 실행
    # "java -jar 실행하는_자바파일"의 형식
    # 여기서 "-Dfile.encoding=UTF-8"은 한글을 깨지지않게 인코딩하는 실행옵션
    # source_name과 result_name을 뒤에 붙혀주어 실행시 자바 main함수의 인자로 넘겨줌

# 3)
move_dir()
execute_jar(modified_osgetcwd + "/" + jar_src_dir)