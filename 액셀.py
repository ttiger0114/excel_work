import numpy as np
import pandas as pd

def main():
    #--------------------파일 경로 설정---------------
    Location = 'C:/Users/ttige/OneDrive/Desktop/'
    File = 'ex.xlsx'
    Output_file_name = 'result.xlsx'
    #-------------------------------------------------


    # 추출 행, 열 선언
    Row = 10000000
    Column = 10000000

    # 추출 및 변환 코드
    data_pd = pd.read_excel('{}/{}'.format(Location, File), 
                     header=None, index_col=None, names=None)
    data_np = pd.DataFrame.to_numpy(data_pd)
    Row = int(data_np.shape[0])
    Column = int(data_np.shape[1])

    # 변경전 출력
    print(data_pd)


    #---------------- 특정 단어 포함시 변경 ----------

    #   A  B  C  D  E  F  G  H  I   K   J   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z
    #   0  1  2  3  4  5  6  7  8   9   10  11  12  13  14  15  16  17  18  19  20  21  22  23  24  25  

    target_column_include = 2;      # 목표 행 설정
    target_include = '파생'
    target_change = '파생형'
    j = 0
    while j < data_np.shape[0]:
        if target_include in str(data_np[j][target_column_include]): # 문자열 부분 일치
            data_np[j][target_column_include] = target_change
        j = j+1
    #---------------------------------------------------




    input = ['!@#$%@#$%!@#' for i in range(Column)]
    #---------------행 별 삭제 타겟 설정-------------

    input[0] = '1245'  # A
    #input[1] = '동동'    # B
    #input[2] = 'fdas'    # C
    #input[3] = '256188'   # D  
    #input[4] = 'DFDF'    # E  
    #input[5] = '4'       # F 
    #input[6] = 'EL       # G
    #input[7] = '5'       # H
    #input[8] = '4'       # I 
    #input[9] = 'ELF'     # J 
    #input[10] = '5'      # K 
    #input[11] = '4'      # L
    #input[12] = 'ELF'    # M
    #input[13] = '5'      # N
    #input[14] = '4'      # O
    #input[15] = 'ELF'    # P
    #input[16] = '5'      # Q 
    #input[17] = '4'      # R
    #input[18] = 'ELF'    # S
    #input[19] = '5'      # T
    #input[20] = '4'      # U
    test_column = [1]
    # 데이터 열 삭제 
    #for i in range(0,Column):
    for i in test_column:
        j = 0
        while j < data_np.shape[0]:
            if input[i] == str(data_np[j][i]) :  # 문자열 완전 일치 
                data_np = np.delete(data_np,j,axis=0)
                j = j-1
            j = j+1


    #------------------------------------------------



    # C~~~ 목표 데이터 column 입력 및 삭제

    #   A  B  C  D  E  F  G  H  I   K   J   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z
    #   0  1  2  3  4  5  6  7  8   9   10  11  12  13  14  15  16  17  18  19  20  21  22  23  24  25  

    target_column = 3       # C가 있는 행
    j = 0
    while j < data_np.shape[0]:
        temp = str(data_np[j][target_column])
        if temp.endswith(('c2','c3','c4','c5','c 2','c 3','c 4','c 5')):
            data_np = np.delete(data_np,j,axis=0)
            j = j-1
        j = j+1

    result = pd.DataFrame(data_np)

    # 변경후 출력
    print(result)

    # 데이터 저장
    direc = Location + Output_file_name
    result.to_excel(direc,index = False)

main() 