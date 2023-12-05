from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# 드라이버 생성
def initialize_driver():
    #서비스 변수 생성
    customService = Service(ChromeDriverManager().install())
    #옵션 변수 생성
    customOption = Options()
    #드라이버 객체 생성
    browser = webdriver.Chrome(service = customService, options = customOption)
    browser.maximize_window() # 창 최대화
    
    return browser

# 아이템 선택
def select_items(browser, items_to_select):
    checkboxes = browser.find_elements(By.NAME, 'fieldIds')

    # 기본으로 체크된 체크 박스 체크 해제
    for checkbox in checkboxes:
        if checkbox.is_selected():    
            checkbox.click()          
    
    # 조회 항목 설정
    for checkbox in checkboxes:
        parent = checkbox.find_element(By.XPATH, '..')
        label = parent.find_element(By.TAG_NAME, 'label')
        if label.text in items_to_select:
            checkbox.click()
    
    # 적용하기 클릭
    btn_apply = browser.find_element(By.XPATH, '//*[@id="contentarea_left"]/div[2]/form/div/div/div/a[1]')# id 같은 게 없으니까 그냥 XPath로 접근하면 편함; XPath가 너무 길면 XPath를 직접 지정해줘도 됨 => '//a[@href="javascript:fieldSubmit()"]: '//'는 html 문서 전체에서 찾겠다는 뜻
    btn_apply.click() 

# 데이터 전처리
def preprocess_data(browser, url):

    for idx in range(1, 50):
        # 페이지 이동
        browser.get(url + str(idx))

        # 데이터 추출
        df = pd.read_html(browser.page_source)[1] # 필요한 테이블만 가져옴
        df.dropna(axis='index', how='all', inplace=True) #html 상에서 구분선 태그 때문에 결측치 발생=>제거!, axis='index' => 행 기준으로 삭제, how='all' => 해당 기준(지금은 행)에 모든 것이 결측치라면 지워라!
        df.dropna(axis='columns', how='all', inplace=True) # 열에도 결측치 있어서 제거(종목토론실이랑 옆에 3개의 불필요한 td 있음)
        df = df.drop('N', axis=1)

        if idx == 1:
            preprocess_df = df
        else:
            preprocess_df = pd.concat([preprocess_df, df])

        if len(df) == 0: # 더 이상 가져올 데이터가 없다면 반복문 탈출
            break    

    return preprocess_df

# 조건에 맞게 데이터 처리
def process_data(df):
    
    #데이터 처리 조건
    condition_for_index = (
        (df.영업이익 > 0) &
        (df.시가총액 > 10000)&
        ((df.보통주배당금 / df.현재가) > 0.03)
    )

    process_df = df.loc[condition_for_index]

    # 조건 만족한 데이터에 배당률 추가
    process_df['배당률'] = round(process_df['보통주배당금'] / process_df['현재가'] * 100, 1).astype(str) + '%'

    return process_df

def select_preferred_stock(df):
    preferred_stock = df[df['종목명'].str.endswith('우')]
    return preferred_stock

def merge_preferred_stock(total_df, preferred_df):

    # priority_data의 '종목명'에서 '우'를 제외한 문자열을 새로운 열에 저장
    preferred_df['종목명_우_제외'] = preferred_df['종목명'].str[:-1]

    # total_data와 priority_data를 '종목명_우_제외'를 기준으로 병합하여 '보통주배당금' 값을 추가
    merged_df = pd.merge(preferred_df, total_df, how='left', left_on='종목명_우_제외', right_on='종목명')
    merged_df = merged_df[['종목명_x', '보통주배당금_y', '영업이익_y']]
    merged_df.rename(columns={'종목명_x': '종목명', '보통주배당금_y': '보통주배당금', '영업이익_y': '영업이익'}, inplace = True)

    # 위 과정을 통해 얻어낸 우선주 '종목명', '영업이익', '보통주배당금' 데이터프레임을 기존 우선주 데이터프레임과 병합
    empty_column_df = merged_df[['종목명', '영업이익', '보통주배당금']]
    merged_preferred_df = pd.merge(preferred_df, empty_column_df, how='inner', on='종목명')
    merged_preferred_df['영업이익_x'] = merged_preferred_df['영업이익_y']
    merged_preferred_df['보통주배당금_x'] = merged_preferred_df['보통주배당금_y']
    merged_preferred_df.rename(columns={'영업이익_x':'영업이익', '보통주배당금_x':'보통주배당금'}, inplace=True)
    merged_preferred_df.drop(['종목명_우_제외','영업이익_y','보통주배당금_y'], axis=1, inplace=True)

    return merged_preferred_df

def main():

    # URL 접속
    browser = initialize_driver()
    URL = 'https://finance.naver.com/sise/sise_market_sum.naver?&page='
    browser.get(URL)

    # 항목 선택 및 결과 출력
    items_to_select = ['시가총액', '영업이익', '보통주배당금']
    select_items(browser, items_to_select)

    # 데이터 처리
    total_df = preprocess_data(browser, URL)
    preferred_stock_df = select_preferred_stock(total_df)
    merged_preferred_df = merge_preferred_stock(total_df, preferred_stock_df)
    matched_merged_preferred_df = process_data(merged_preferred_df)
    matched_total_df = process_data(total_df)
    result_df = pd.concat([matched_total_df, matched_merged_preferred_df])

    # csv로 저장
    result_df.to_csv('final.csv', encoding='utf-8-sig', index=False) 

if __name__ == "__main__":
    main()