import os
from datetime import datetime

def read_file_to_variable(file_path, encoding='utf-8'):
    """
    특정 경로의 파일을 읽어 지정된 인코딩으로 변수에 할당합니다.

    :param file_path: 파일 경로 (str)
    :param encoding: 파일 인코딩 (str, 기본값 'utf-8')
    :return: 파일 내용 (str)
    """
    try:
        with open(file_path, 'r', encoding=encoding) as file:
            content = file.read()
        return content
    except Exception as e:
        print(f"An error occurred while reading the file: {e}")
        return None


def delete_kakaotalk_files(directory_path):
    """
    지정된 디렉토리에서 'KakaoTalk'로 시작하는 모든 파일을 삭제합니다.

    :param directory_path: 디렉토리 경로 (str)
    """
    try:
        if not os.path.exists(directory_path):
            print(f"Directory does not exist: {directory_path}")
            return

        for filename in os.listdir(directory_path):
            if filename.startswith("KakaoTalk"):
                file_path = os.path.join(directory_path, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
    except Exception as e:
        print(f"An error occurred while deleting files: {e}")


def filter_text_lines(kakao_opentalk_name, input_file_path, output_file_path):
    """
    입력 파일에서 '['로 시작하는 행을 제외한 모든 텍스트를 삭제하고 결과를 새로운 파일에 저장합니다.

    :param input_file_path: 원본 파일 경로 (str)
    :param output_file_path: 출력 파일 경로 (str)
    """
    print("input_file_path : ", input_file_path)

    try:
        # 출력 디렉토리 생성
        output_dir = os.path.dirname(output_file_path)
        print("output_dir : ", output_dir)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 현재 날짜와 시간을 파일명에 추가
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_file_name = f"{kakao_opentalk_name}_{timestamp}.txt"
        new_output_file_path = os.path.join(output_dir, new_file_name)

        print("New output file path: ", new_output_file_path)

        # 파일 읽기
        with open(input_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        # '['로 시작하는 행만 필터링
        filtered_lines = [line for line in lines if line.lstrip().startswith('[')]

        # 결과를 새로운 파일에 저장
        with open(new_output_file_path, 'w', encoding='utf-8') as file:
            file.writelines(filtered_lines)

        print(f"Filtered text saved to {new_output_file_path}")
        return new_output_file_path
    except Exception as e:
        print(f"An error occurred: {e}")
        return e