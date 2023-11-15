import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QToolTip, QMessageBox, QLineEdit
from PyQt5.QtGui import QIcon, QPixmap, QFont
from PyQt5.QtCore import QCoreApplication, QDateTime, QTimer
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd


# 주식 데이터를 네이버에서 크롤링 하여 확인한다.
class jusikplay(QWidget):
    def __init__(self):
        super().__init__()
        self.uiInit()
        self.reset_timer = QTimer(self)
        self.reset_timer.timeout.connect(self.enableResetButton)

    #UI 초기화 메서드
    def uiInit(self):
        # NAVER 국내주식의 타이틀 앞부분 UI 라벨 
        self.titleLabelOne = QLabel('NAVER', self)
        self.titleLabelOne.move(7, 25)
        self.titleLabelOne.setFont(QFont('Arial',15,85))
        self.titleLabelOne.setStyleSheet('color: rgb(0, 204, 102)')
        
        # NAVER 국내주식의 타이틀 뒷부분 UI 라벨
        self.titleLabelTwo = QLabel('국내주식', self)
        self.titleLabelTwo.move(96, 25)
        self.titleLabelTwo.setFont(QFont('Arial',12,85))

        # 당일날짜 요일 라벨
        self.todayLabel = QLabel(self.checkdate(),self)
        self.todayLabel.move(430, 25)
        self.todayLabel.setFont(QFont('Arial',12,85))


        # 코스피, 코스닥, 코스피200 top10 인기종목순위 보여주기 
        kospiNow, kospiChange, kosdaqNow, kosdaqChange, kpi200Now, kpi200Change, top10List = self.getBasicInfoAndRankTop10()
        
        # 코스피 정보 라벨
        self.basicKospi = QLabel('코스피 : ', self)
        self.basicKospi.move(10, 75)
        self.basicKospi.setFont(QFont('Arial',12,60))

        self.basicKospiValue = QLabel(f'{kospiNow}', self)
        self.basicKospiValue.move(90, 65)
        self.basicKospiValue.setFont(QFont('Arial',10,60))

        self.basicKospiChange = QLabel(f'{kospiChange}',self)
        self.basicKospiChange.move(90, 70)
        self.basicKospiChange.setFont(QFont('Arial',10,60))

        #코스닥 정보 라벨
        self.basicKosda = QLabel('코스닥 : ', self)
        self.basicKosda.move(205, 75)
        self.basicKosda.setFont(QFont('Arial',12,60))

        self.basicKosdaValue = QLabel(f'{kosdaqNow}', self)
        self.basicKosdaValue.move(290, 65)
        self.basicKosdaValue.setFont(QFont('Arial',10,60))

        self.basicKospiChange = QLabel(f'{kosdaqChange}',self)
        self.basicKospiChange.move(290, 70)
        self.basicKospiChange.setFont(QFont('Arial',10,60))

        #코스피200 정보 라벨
        self.basicKospi200 = QLabel('코스피200 : ', self)
        self.basicKospi200.move(400, 75)
        self.basicKospi200.setFont(QFont('Arial',12,60))

        self.basicKospi200Value = QLabel(f'{kpi200Now}', self)
        self.basicKospi200Value.move(515, 65)
        self.basicKospi200Value.setFont(QFont('Arial',10,60))

        self.basicKospi200Change = QLabel(f'{kpi200Change}',self)
        self.basicKospi200Change.move(515, 70)
        self.basicKospi200Change.setFont(QFont('Arial',10,60))

        # 인기검색종목 라벨
        self.ranktop10Label = QLabel('인기종목 Top 10', self)
        self.ranktop10Label.move(20, 130)
        self.ranktop10Label.setFont(QFont('Arial',14,60))
        self.ranktop10Label.setStyleSheet('color: rgb(0, 0, 255)')

        # 10개의 인기종목을 가져와서 보여주기 위한 작업
        for i, stock_info in enumerate(top10List): 
            ranklabel = QLabel(stock_info, self)
            ranklabel.move(20, 170 + i * 30)  # 라벨을 순차적으로 배치
            ranklabel.setFont(QFont('Arial', 10, 60))

        # 기본정보 새로고침 버튼
        self.resetbasicbutton = QPushButton('새로고침', self)
        self.resetbasicbutton.move(195, 124)
        self.resetbasicbutton.resize(80, 40)
        self.resetbasicbutton.clicked.connect(self.resetBasicInfo)
        self.resetbasicbutton.setToolTip('30초에 1회 새로고침이 가능합니다.')


        # 종목 검색 라벨
        self.searchstock = QLabel('종목명 : ',self)
        self.searchstock.move(350, 130)
        self.searchstock.setFont(QFont('Arial',12,60))

        self.line_edit = QLineEdit(self)
        self.line_edit.move(415,125)
        self.line_edit.resize(140,35)
        self.line_edit.setToolTip('종목명을 입력해주세요 (띄어쓰기X)')

        self.searchButton = QPushButton('검색', self)
        self.searchButton.move(560, 125)
        self.searchButton.resize(80, 35)
        self.searchButton.clicked.connect(self.searchStockInfo)

        # 종목 검색 결과값
        # 종목의 현재 상태 
        self.nowInfo = QLabel('현재 : ',self)
        self.nowInfo.move(330, 170)
        self.nowInfo.setFont(QFont('Arial',10,60))
        self.nowInfo.hide()

        # 종목 실시간 가격정보
        self.nowPriceLabel = QLabel('현재가 : ',self)
        self.nowPriceLabel.move(330, 200)
        self.nowPriceLabel.setFont(QFont('Arial',10,60))
        self.nowPriceLabel.resize(400,30)
        self.nowPriceLabel.hide()
        
        self.nowPriceInfoLbel = QLabel(' : ',self)
        self.nowPriceInfoLbel.move(330, 230)
        self.nowPriceInfoLbel.setFont(QFont('Arial',10,60))
        self.nowPriceInfoLbel.resize(400,30)
        self.nowPriceInfoLbel.setStyleSheet("background-color: lightgray;")
        self.nowPriceInfoLbel.hide()
        
        #종목 거래정보
        self.yesterdadyPrice = QLabel('전일 : ',self)
        self.yesterdadyPrice.move(330, 260)
        self.yesterdadyPrice.setFont(QFont('Arial',10,60))
        self.yesterdadyPrice.resize(400,30)
        self.yesterdadyPrice.hide()       

        self.highPrice = QLabel('고가 : ',self)
        self.highPrice.move(330, 290)
        self.highPrice.setFont(QFont('Arial',10,60))
        self.highPrice.resize(400,30)
        self.highPrice.hide()  
        
        self.upperLimit = QLabel('상한가 : ',self)
        self.upperLimit.move(490, 290)
        self.upperLimit.setFont(QFont('Arial',10,60))
        self.upperLimit.resize(400,30)
        self.upperLimit.hide()          
        
        self.volumeLabel = QLabel('거래량 : ',self)
        self.volumeLabel.move(330, 320)
        self.volumeLabel.setFont(QFont('Arial',10,60))
        self.volumeLabel.resize(400,30)
        self.volumeLabel.hide()  

        self.startPrice = QLabel('시가 : ',self)
        self.startPrice.move(330, 350)
        self.startPrice.setFont(QFont('Arial',10,60))
        self.startPrice.resize(400,30)
        self.startPrice.hide()  

        self.lowPrice = QLabel('저가 : ',self)
        self.lowPrice.move(330, 380)
        self.lowPrice.setFont(QFont('Arial',10,60))
        self.lowPrice.resize(400,30)
        self.lowPrice.hide()  

        self.lowerLimit  = QLabel('하한가 : ',self)
        self.lowerLimit.move(490, 380)
        self.lowerLimit.setFont(QFont('Arial',10,60))
        self.lowerLimit.resize(400,30)
        self.lowerLimit.hide()  

        self.transactionPrice  = QLabel('거래대금 : ',self)
        self.transactionPrice.move(330, 410)
        self.transactionPrice.setFont(QFont('Arial',10,60))
        self.transactionPrice.resize(400,30)
        self.transactionPrice.hide()

        # 버튼 
        
        #엑셀 다운로드 버튼
        self.excelDwBtn = QPushButton('엑셀다운로드', self)
        self.excelDwBtn.move(130, 500)
        self.excelDwBtn.resize(200, 50)
        self.excelDwBtn.clicked.connect(self.dwExcel)
        self.excelDwBtn.setToolTip('모든종목 데이터를 다운로드 합니다.')


        #프로그램 종료 버튼
        self.exitProgramBtn = QPushButton('프로그램종료', self)
        self.exitProgramBtn.move(370, 500)
        self.exitProgramBtn.resize(200, 50)
        self.exitProgramBtn.clicked.connect(self.exitProgram)
        self.exitProgramBtn.setToolTip('프로그램을 종료합니다.')

        # 프로그램 창 타이틀명
        self.setWindowTitle('NAVER 주식보기')
        self.setWindowIcon(QIcon('img/naverjusicicon.jpg'))
        self.setGeometry(600, 300, 660, 560)
        self.show()
        

    # 날자 요일 구하기 메서드
    # date.time.weekday를 활용하여 요일 리스트에서 요일을 사용하고 현재 날짜값을 구함
    def checkdate(self):
        koreaDays = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일','일요일']
        nowDateTime = datetime.now()
        dayToNumber = nowDateTime.weekday()
        koreaDay = koreaDays[dayToNumber]
        
        formatNowDateTime = nowDateTime.strftime(f'%Y.%m.%d {koreaDay}')
        return formatNowDateTime
        

    # 새로고침 버튼 메서드
    # 30초에 1회 가능하도록 작업
    def resetBasicInfo(self):
        self.getBasicInfoAndRankTop10() # 실제 네이버 주식의 검색종목 순의 10개를 가져오는 함수
        self.resetbasicbutton.setEnabled(False) 
        self.reset_timer.start(30000) # 30초 동안 버튼 잠금 
        
    # 새로고침 버튼 초기화 메서드  
    def enableResetButton(self):
        self.resetbasicbutton.setEnabled(True)
        self.reset_timer.stop()

    # 엑셀 다운로드 메서드
    def dwExcel(self):
        try:
            QMessageBox.about(self, '진행', '엑셀 다운로드를 진행합니다. 10초 정도 소요됩니다.')
            resultDf = pd.DataFrame(columns=['순위', '종목명', '현재가', '전일비', '등략률', '액면가', '시가총액', '상장주식수', '외국인비율', '거래량', 'PER', 'ROE', '토론실'])

            for page in range(1, 41):
                url = f'https://finance.naver.com/sise/sise_market_sum.naver?sosok=0&page={page}'  # 페이지 번호를 포함한 URL

                res = requests.get(url)
                res.encoding = 'euc-kr'
                res.raise_for_status()  # 네트워크 오류 발생시 예외 발생
                soup = BeautifulSoup(res.text, 'html.parser')

                table = soup.find('table', class_='type_2')

                data = [] # 엑셀에 입력할 데이터를 담을 빈 리스트
                
                # table 에서 TR 만 전부 가져와 FOR 문을 돌고 그 안에서 또 TD 만 다 가져와서 필요한 값만 공백제거하고 DATA에 넣는다  
                for row in table.find_all('tr'):
                    cols = [col.text.strip() for col in row.find_all('td')]
                    if cols: # 열데이터 공백 여부 확인 후 넣기.
                        data.append(cols)            
                            
                            
                            
                # 엑셀에 데이터 넣기
                tempDf = pd.DataFrame(data, columns=resultDf.columns)
                resultDf = pd.concat([resultDf, tempDf])

            resultDf = resultDf.dropna()
            saveData = datetime.now()
            saveData = saveData.strftime('%Y.%m.%d %H_%M_%S') #파일명에 저장하는 날짜 시간을 표시하여 구분하기 위함

            # 엑셀 파일 저장 시 예외 처리
            try:
                resultDf.to_excel(f'{saveData}_stock_data.xlsx', index=False)
                QMessageBox.about(self, '완료', '엑셀 다운로드 완료.')
            except Exception as e:
                QMessageBox.critical(self, '파일 저장 오류', f'엑셀 파일을 저장하는 중 오류가 발생했습니다: {str(e)}')

        except requests.RequestException as e:
            QMessageBox.critical(self, '네트워크 오류', f'네트워크 연결에 문제가 있습니다: {str(e)}')

        
    
    #프로그램 종료 버튼 메서드
    def exitProgram(self):
        return QCoreApplication.instance().quit()

    # 기초 UI화면의 데이터를 보여줄 크롤링 메서드
    # 실제 네이버 주식의 검색종목 순의 10개를 가져오는 함수
    def getBasicInfoAndRankTop10(self):
        try:
            url = 'https://finance.naver.com/sise/'
            res = requests.get(url)
            res.raise_for_status() 
            res.encoding = 'euc-kr'
            html = res.text
            soup = BeautifulSoup(html, 'html.parser')

            # 기본 정보 가져오기
            kospiNow = soup.select('#KOSPI_now')[0].text
            kospiChange = soup.select('#KOSPI_change')[0].text.replace('상승', '')

            kosdaqNow = soup.select('#KOSDAQ_now')[0].text
            kosdaqChange = soup.select('#KOSDAQ_change')[0].text.replace('상승', '')

            kpi200Now = soup.select('#KPI200_now')[0].text
            kpi200Change = soup.select('#KPI200_change')[0].text.replace('상승', '')

            # 인기 검색 종목 가져오기
            stock_elements = soup.select('.lst_pop li a, .lst_pop li a + span.up, .lst_pop li a + span.dn')

            top10StockList = []  # 상위 10개 주식 정보를 저장할 리스트

            stock_info = []  # 각 주식 정보를 임시로 저장할 리스트
            for element in stock_elements:
                stock_info.append(element.text)

                if len(stock_info) == 2:
                    top10StockList.append(f'{stock_info[0]} | {stock_info[1]}')
                    stock_info = []

            return kospiNow, kospiChange, kosdaqNow, kosdaqChange, kpi200Now, kpi200Change, top10StockList

        except requests.RequestException as e:
            QMessageBox.critical(self, '네트워크 오류', f'네트워크 연결에 문제가 있습니다: {str(e)}')
        except Exception as e:
            QMessageBox.critical(self, '오류', f'데이터 가져오는 중 오류가 발생했습니다: {str(e)}')

    # 검색하는 종목의 데이터를 크롤링 하는 메서드드
    def searchStockInfo(self):
        searchkeyword = self.line_edit.text()

        try:
            df = pd.read_excel('./dt/종목코드.xlsx', dtype={'종목코드': str})

            filteredRow = df[df['회사명'] == searchkeyword]

            if not filteredRow.empty:
                keywordCode = filteredRow.at[filteredRow.index[0], '종목코드']
                
                url = 'https://finance.naver.com/item/main.naver?code='
                try:
                    searchUrl = url + keywordCode
                    res = requests.get(searchUrl)
                    res.encoding = 'euc-kr'
                    html = res.text
                    soup = BeautifulSoup(html, 'html.parser')
                    
                    #날짜 기준 상태
                    dateEmTag = soup.find('em', class_='date')
                    dateText = dateEmTag.get_text(strip=True)
            

                    # 현재가격
                    nowPriceTag = soup.select_one('.no_today')
                    nowPrice = nowPriceTag.find('span', class_='blind').text 

                    
                    # 현재 상태 전일대비 얼마나 상승 얼마나 하락 정보
                    nowStateTag = soup.find('p', class_='no_exday')
                    if nowStateTag:
                        spans = nowStateTag.find_all('span', class_='sptxt sp_txt1') + nowStateTag.find_all('span', class_='ico') + nowStateTag.find_all('span', class_='blind')

                        nowState = [] 
                        for span in spans:
                            nowState.append(span.text) 
                        
                        # nowState = [span.text for span in spans]

                        nowStateList = ' '.join(nowState)
                        nowStateListResult = nowStateList.split()

                    else:
                        nowStateListResult =[]
                    
                    #시가 고가 저가 거래량 등 정보
                    #todo 진짜 naver 퍼블리셔가 일부로 이렇게 만든건가 금액 단위 다 태그로 쪼개서... 혹시 모르기 이것도 검색해서 코드 수정필요
                    previous_closing_prices = soup.select('td span.sptxt.sp_txt2 + em span[class^="no"]')
                    high_prices = soup.select('td span.sptxt.sp_txt4 + em span[class^="no"]')
                    low_prices = soup.select('td span.sptxt.sp_txt5 + em span[class^="no"]')
                    volumes = soup.select('td span.sptxt.sp_txt9 + em span[class^="no"]')
                    start_prices = soup.select('td span.sptxt.sp_txt3 + em span[class^="no"]')
                    uper_limits = soup.select('td span.sptxt.sp_txt6 + em span[class^="no"]')
                    lower_limits = soup.select('td span.sptxt.sp_txt7 + em span[class^="no"]')
                    trade_amounts = soup.select('td span.sptxt.sp_txt10 + em span[class^="no"]')


                    # 태그요소의 값을 받아서 텍스트들만 추출하고 공백을 지운뒤 숫자로 리턴한다.
                    # 크롤링할때 태그가 다 쪼개저 있어서 1,400 처럼 한화화폐 표현이 불가능하여 작성.
                    def get_text_with_no(elements):
                        text = ''.join([element.text for element in elements])
                        return int(text)

                    # format을 사용해서 숫자 값을 , 한화화폐 표현으로 변경 
                    yester = format(get_text_with_no(previous_closing_prices),',')
                    highPr = format(get_text_with_no(high_prices),',')
                    lowPr = format(get_text_with_no(low_prices),',')
                    volume = format(get_text_with_no(volumes),',')
                    startPr = format(get_text_with_no(start_prices),',')
                    upppr = format(get_text_with_no(uper_limits),',')
                    lowerPr = format(get_text_with_no(lower_limits),',')
                    transactionPr = format(get_text_with_no(trade_amounts),',')

                    self.nowInfo.setText(f'현재 : {dateText}')
                    self.nowInfo.show()
                    
                    self.nowPriceLabel.setText(f'현재가 : {nowPrice}')
                    self.nowPriceLabel.show()
                    
                    # 상승 하락인경우는 리스트 크기가 같았는데 보합인경우에는 달라서 따로 처리필요
                    if len(nowStateListResult) > 4 :
                        self.nowPriceInfoLbel.setText(f'{nowStateListResult[0]} {nowStateListResult[2]} {nowStateListResult[3]} {nowStateListResult[1]} {nowStateListResult[2]} {nowStateListResult[4]}%')
                        self.nowPriceInfoLbel.show()
                    else:
                        self.nowPriceInfoLbel.setText(f'{nowStateListResult[0]}  {nowStateListResult[2]}  {nowStateListResult[1]} {nowStateListResult[3]}%')
                        self.nowPriceInfoLbel.show()

                    self.yesterdadyPrice.setText(f'전일 : {yester}원')
                    self.yesterdadyPrice.show()                

                    self.highPrice.setText(f'고가 : {highPr}원')
                    self.highPrice.show()

                    self.upperLimit.setText(f'상한가 : {upppr}원 ')
                    self.upperLimit.show()   

                    self.volumeLabel.setText(f'거래량 : {volume}')
                    self.volumeLabel.show()                
                    
                    self.startPrice.setText(f'시가 : {startPr}원')
                    self.startPrice.show()               
                    
                    self.lowPrice.setText(f'저가 : {lowPr}원')
                    self.lowPrice.show()               
                    
                    self.lowerLimit.setText(f'하한가 : {lowerPr}원')
                    self.lowerLimit.show()
                    
                    self.transactionPrice.setText(f'거래대금 : {transactionPr}백만')
                    self.transactionPrice.show()
                    
                except requests.exceptions.RequestException as e:
                    QMessageBox.critical(self, '서버 접속 오류', f'서버에 접속 중 오류가 발생했습니다: {e}')                
            else:
                QMessageBox.warning(self, '검색 실패', f"'{searchkeyword}'에 해당하는 종목을 찾을 수 없습니다.")
        except FileNotFoundError:
            QMessageBox.critical(self, '파일 오류', '/dt/종목코드.xlsx 파일을 찾을 수 없습니다.')


if __name__ == '__main__':
    showplay = QApplication(sys.argv)
    inginstens = jusikplay()
    showplay.exec_()
