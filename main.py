import pandas as pd
import numpy as np

#파일을 불러서 데이터프레임을 작성하고 헤더 전처리 function
def read_excel(file_direction):
    df = pd.read_excel(file_direction).applymap(str) #헤더 확인을 위해 엑셀 1 불러옴, Value는 문자열로 변환
    df_before = df.applymap(str.upper).applymap(lambda x: pd.to_numeric(x, errors='ignore')).replace('NAN',np.NaN) # Value 전체 대문자로 변환, 숫자형 문자는 숫자로 재변환, NAN문자는 NaN으로 변경
    df_before.columns = [str(x).upper() for x in df_before.columns] #칼럼이 int, str 혼용되어 있을 경우 str은 대문자 처리하고 int는 그대로 둠
    ref_tags = ['TAG NO.','FULL TAG NO.','TAG NUMBER'] #TAG NO. 칼럼으로 인식할 칼럼명 기입
    break_flag = False #break_flag = False #이중 반복문 탈출용 flag

    #TAG NO. 칼럼의 행이 1행에 없어도 TAG NO.를 dataframe의 칼럼으로 인식 할 수 있게함
    for column_tag in df_before.columns:  # 불러온 파일의 칼럼 전체 조사
        if break_flag == True: #아래 if문의 else가 동작하면, break 발동하여 for 문 최종 탈출
            break
        for ref_tag in ref_tags: #기준 칼럼명과 읽어오는 파일의 칼럼 하나씩 비교
            if column_tag == ref_tag: #불러온 칼럼에 기준 칼럼명이 하나라도 있을 경우
                break_flag = True  # 2중 for 문 탈출용
                return df_before # 1행에 기준 칼럼이 있으므로, 반복문을 종료하고 읽어온 파일을 return한다.
                break_flag = True  # 2중 for 문 탈출용
                break
            else:
                try:
                    column_index = int(df_before[df_before[column_tag] == ref_tag].index.values)
                    df_temp = pd.read_excel(file_direction, header=column_index + 1).rename(columns=str.upper).applymap(
                        str)  # 추출한 열번호 기준으로 헤더 설정하고 파일 읽어옴
                    return df_temp.applymap(str.upper).applymap(lambda x: pd.to_numeric(x, errors='ignore')).replace(
                        'NAN', np.NaN)  # Value 전체 대문자로 변환, 숫자형 문자는 숫자로 재변환, NAN문자는 NaN으로 변경
                    break_flag = True  # 2중 for 문 탈출용
                    break
                except:
                    error_flag = True

#a = read_excel('data_3d_real.xlsx')

'''
    for i in range(0, len(tag_no), 1):
        if df_before.columns.any() != tag_no[i]:  # TAG NO. 칼럼이 1행에 없을 경우 아래 진행
            for j in df_before.columns: #개별 칼럼의 TAG NO. 행 번호 확인
                if df_before[df_before[j]==tag_no[i]].index.values.size > 0: #TAG NO.가 있는 열번호 반환
                    header_no_1 = int(df_before[df_before[j]==tag_no[i]].index.values)
                    df_temp = pd.read_excel(file_direction, header=header_no_1 + 1).rename(columns=str.upper).applymap(str)  # 추출한 열번호 기준으로 헤더 설정하고 파일 읽어옴
                    return df_temp.applymap(str.upper).applymap(lambda x: pd.to_numeric(x, errors='ignore')).replace('NAN',np.NaN) # Value 전체 대문자로 변환, 숫자형 문자는 숫자로 재변환, NAN문자는 NaN으로 변경
        else:
            return df_before
'''