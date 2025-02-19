# pyinstaller -w -F --uac-admin --add-data "static;static" --icon=.\static\icon.ico --noconfirm marginguide_explorer.py
from seleniumbase import SB
import time, os, random,sqlite3,winsound, sys
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import tkinter as tk
from threading import Timer
import pandas as pd
now = datetime.today()
str_today = str(datetime.strftime(now, "%Y-%m-%d"))
basedir = os.path.abspath(os.path.dirname(__file__))
sel_path = os.path.join(basedir, "data", "seldb.db")
db_path = os.path.join(basedir, "data", "db.db")
# db_path = "C:\\Program Files (x86)\\Margin Guide\\data\\db.db"
# sel_path = "C:\\Program Files (x86)\\Margin Guide\\data\\seldb.db"
try:
    try:
        # 마진가이드용 DB
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        # 셀레니움용 DB
        con = sqlite3.connect(sel_path)
        cur = con.cursor()
    except:
        conn = sqlite3.connect('data/db.db')
        cursor = conn.cursor()
        con = sqlite3.connect('data/seldb.db')
        cur = con.cursor()
except Exception as e:
    pass


    
def get_sound_path(filename):
    """PyInstaller 실행 환경에서도 올바른 경로를 반환하는 함수"""
    if getattr(sys, 'frozen', False):  # .exe 실행 여부 확인
        base_path = os.path.join(sys._MEIPASS, "static")  # static 폴더 포함
    else:
        base_path = os.path.abspath("static")  # 개발 환경에서는 일반적인 폴더 사용
    return os.path.join(base_path, filename)

def show_custom_notification_(data_1 = '', data_2 = ''):
    sound_path = get_sound_path("ranking.wav")
    # 알림 소리 재생
    if os.path.exists(sound_path):
        winsound.PlaySound(sound_path, winsound.SND_ASYNC)
    try:
        # 알림 창 생성
        root = tk.Tk()
        root.overrideredirect(True)  # 제목 표시줄 숨김
        root.attributes("-topmost", True)  # 항상 위에 표시
        root.attributes("-alpha", 0.9)  # 반투명 효과

        # 화면 크기 가져오기
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 창 크기 및 위치 설정
        window_width = 360
        window_height = 200
        x_position = screen_width - window_width - 20
        y_position = screen_height - window_height - 50
        root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # 배경 설정
        root.configure(bg="#1E1E1E")  # 다크 그레이 배경

        # 내부 프레임
        frame = tk.Frame(root, bg="#2A2A2A", relief="flat", bd=2)
        frame.place(relx=0.5, rely=0.5, anchor="center", width=340, height=200)

        # 닫기 버튼 (우측 상단 X 버튼)
        def close_app():
            try:
                root.destroy()
            except:
                pass
            os._exit(0)

        close_button = tk.Button(frame, text="❌", command=close_app, bg="#2A2A2A", fg="#BBBBBB", font=("맑은 고딕", 10, "bold"),
                                relief="flat", bd=0, padx=5, pady=2, activebackground="#444444", activeforeground="white")
        close_button.place(x=305, y=10)  # 우측 상단 배치

        # 제목 레이블
        title_label = tk.Label(frame, text="📊 MarginGuide", font=("맑은 고딕", 15, "bold"), fg="#FFFFFF", bg="#2A2A2A")
        title_label.pack(pady=(15, 25))

        # 메시지 텍스트
        msg_label = tk.Label(frame, text="검색어 랭킹 알림", font=("맑은 고딕", 11), bg="#2A2A2A", fg="#DDDDDD")
        msg_label.pack(pady=(0, 10))

        # 데이터 표시 (grid 정렬)
        data_frame = tk.Frame(frame, bg="#2A2A2A")
        data_frame.pack(pady=5)

        # 아이콘과 텍스트 정렬을 맞추기 위해 grid 사용
        text_size = ("맑은 고딕", 13, "bold")

        tk.Label(data_frame, text=data_1, font=text_size, fg="#EEEEEE", bg="#2A2A2A").grid(row=0, column=1, sticky="w")
        tk.Label(data_frame, text=data_2, font=text_size, fg="#EEEEEE", bg="#2A2A2A").grid(row=1, column=1, sticky="w")

        # 페이드 인 효과
        root.attributes("-alpha", 0.0)  # 처음에는 투명

        def fade_in(opacity=0.0):
            if opacity < 0.9:
                opacity += 0.1
                root.attributes("-alpha", opacity)
                root.after(50, lambda: fade_in(opacity))

        fade_in()  # 실행

        # 일정 시간이 지나면 창 닫기
        def timeout():
            os._exit(0)
        Timer(10, timeout).start()  # 10시간 후 자동 종료간 후 자동 종료

        root.mainloop()
    except:
        pass

def input_log(content):
    err_time = datetime.now()
    err_time = datetime.strftime(err_time, "%Y-%m-%d %H:%M:%S")
    query = f"INSERT INTO log (account , actionname, context, logdate) VALUES ('공용', '순위검색', '{content}', '{err_time}')"
    cursor.execute(query)
    conn.commit()

def keyword_to_rank():
    # 키워드 목록 가져오기
    query = f"SELECT keyword FROM search_keyword WHERE last_check != '{str_today}' OR last_check IS NULL"
    df = pd.read_sql_query(query, conn)
    keywords = df['keyword'].tolist()
    return keywords

def update_date(keyword):
    query = f"UPDATE search_keyword SET last_check = '{str_today}' WHERE keyword = '{keyword}'"
    cursor.execute(query)
    conn.commit()
    return True

def insert_endtime(till):
    query = f"""INSERT INTO use_selenium (process, last_time) 
                VALUES ( 'search_rank', '{till}') 
                ON CONFLICT DO UPDATE SET 
                last_time = '{till}'
                """
    cursor.execute(query)
    conn.commit()
    return True

def delete_endtime():
    try:
        query = f"""UPDATE use_selenium SET last_time = '' WHERE process = 'search_rank'
                    """
        cursor.execute(query)
        conn.commit()
        return True
    except:
        return True

def is_headless():
    try:
        query =  "SELECT set_value FROM setting WHERE setting_name = 'headless'"
        df = pd.read_sql_query(query, conn)
        result = df['set_value'][0]
        if result == "on":
            data = True
        else:
            data = False
    except:
        data = True
    return data


    
    
def rankings():
    try:
        coupang_url =f'https://www.coupang.com' 
        result = []
        # 내 상품의 optcode list
        query = "SELECT optcode, prdcode FROM optlist"
        df = pd.read_sql_query(query, conn)
        my_opt = df.set_index(str("optcode"))[str("prdcode")].to_dict()
        input_log('p-1')
        keywords = keyword_to_rank()
        input_log('p-2')

        if len(keywords) > 0:
            # 시간계산
            input_log('p-3')
            op_time = now + timedelta(seconds=230 * len(keywords))
            input_log('p-4')
            op_time = datetime.strftime(op_time, "%Y-%m-%d %H:%M:%S")
            input_log('p-5')
            insert_endtime(op_time)
            input_log('p-6')
            headless = is_headless()
            input_log('p-7')
            with SB(headless2=headless, uc=True, log_cdp=True, block_images=True,undetectable=True,incognito=False) as self:
                self.open('http://google.com')
                success_cnt = 0
                for keyword in keywords:
                    
                    optcode_list = []
                    rankings = []
                    rank = 0
                    page = 1
                    url = f"https://www.coupang.com/np/search?rocketAll=false&searchId=e6a8eb393372964&q={keyword}&brand=&offerCondition=&filter=&availableDeliveryFilter=&filterType=&isPriceRange=false&priceRange=&minPrice=&maxPrice=&page={page}&trcid=&traid=&filterSetByUser=true&channel=user&backgroundColor=&component=&rating=0&sorter=scoreDesc&listSize=72"
                    self.open(url)
                    if self.is_text_visible("ERR_HTTP2_PROTOCOL_ERROR"):
                        input_log('ERR_HTTP2_PROTOCOL_ERROR')
                        show_custom_notification_(data_1 = 'ERR_HTTP2_PROTOCOL_ERROR', data_2=  '관리자에게 문의해주세요.')
                        os._exit()
                        return False
                    
                    if self.is_text_visible("Access_Denied"):
                        input_log('Access_Denied')
                        show_custom_notification_(data_1 = '쿠팡사이트가 검색을 차단했습니다.', data_2=  '30분 이후 다시 시도해주세요.')
                        os._exit()
                        return False
                    try:last_page = int(self.get_text("//a[contains(@class, 'btn-last')]", timeout=4))
                    except:last_page = len(self.find_elements("//span[@class='btn-page']/a"))
                    while True:

                        if self.is_text_visible("Access Denied"):
                            input_log('Access_Denied')
                            show_custom_notification_(data_1 = '쿠팡사이트가 검색을 차단했습니다.', data_2=  '30분 이후 다시 시도해주세요.')
                            os._exit()
                            return False
                        time.sleep(random.randint(3000, 5000) / 1000)
                        scroll = random.randint(10, 1000) * 10
                        self.execute_script(f"window.scrollTo(0, {scroll})")
                        time.sleep(random.randint(1000, 3000) / 1000)
                        soup = BeautifulSoup(self.get_page_source(), "html.parser")
                        items = soup.select("""
                                            li.search-product:not(.sdw-aging-carousel-item):not(.ad-badge-text):not(.search-product__ad-badge)
                                            """)
                        disp_cnt = 0
                        for index, item in enumerate(items, start=1): #저장한 items들을 순차적으로 검색함.  
                            disp_cnt += 1
                            mine = 0
                            # 광고면 지나가
                            try: 
                                ads = item.find('span', class_='ad-badge-text').get_text(strip=True) 
                                continue
                            except:
                                pass
                            
                            # 시간 특가 등 광고이면 지나가
                            try: 
                                ads = item.find('span', class_='sdw-aging-carousel-item').get_text(strip=True) 
                                continue
                            except:
                                pass

                            rank += 1
                            try:optcode = item['data-vendor-item-id']
                            except:continue

                            if optcode in optcode_list:
                                continue
                            else:
                                optcode_list.append(optcode)
                            prdcode = my_opt.get(optcode, '')

                            if rank  > 70 and prdcode == '':
                                continue

                            if optcode in my_opt:
                                mine = 1
                            try:dispcode = item['data-product-id']
                            except:continue
                            try:dispname = item.find('div', class_='name').get_text(strip=True)
                            except:continue
                            try:
                                price = item.find('strong', class_='price-value').get_text(strip=True)
                                price = price.replace(",", "")
                                price = int(price)
                            except:continue
                            try:
                                thumb = item.find('img', class_='search-product-wrap-img')['src']
                                n = thumb.find('/image/')
                                thumb = "https://thumbnail10.coupangcdn.com/thumbnails/remote/120x120" + thumb[n:]
                            except:continue
                            try:
                                link = item.find('a')['href']
                                link = coupang_url + link
                            except:continue
                            try: 
                                rating = item.find('em', class_='rating').get_text(strip=True)
                                rating = float(rating)
                            except: rating = 0
                            
                            try: 
                                rating_cnt = item.find('span', class_='rating-total-count').get_text(strip=True).replace('(', '').replace(')', '')
                                rating_cnt = int(rating_cnt)
                            except: rating_cnt = 0
                            try:
                                deltype = item.find('img', alt='로켓배송')['src']
                                if "Merchant" in deltype:
                                    deltype = "제트배송"
                                elif "global" in deltype:
                                    deltype = "로켓직구"
                                else:
                                    deltype = "로켓배송"
                            except:
                                deltype = "판매자배송"
                            disp_page = int(page) * 2
                            if disp_cnt <= 36:
                                disp_page = disp_page- 1
                            rankings.append((
                                    str_today,  keyword, disp_page, deltype, rank, prdcode, optcode, dispcode,  dispname,  price, thumb, link, rating, rating_cnt, mine
                                ))
                        page += 1
                        if page > last_page:
                            break
                        try:
                            xpath = "//a[@class='btn-next']"
                            self.click(xpath)
                        except:
                            url = f"https://www.coupang.com/np/search?rocketAll=false&searchId=e6a8eb393372964&q={keyword}&brand=&offerCondition=&filter=&availableDeliveryFilter=&filterType=&isPriceRange=false&priceRange=&minPrice=&maxPrice=&page={page}&trcid=&traid=&filterSetByUser=true&channel=user&backgroundColor=&component=&rating=0&sorter=scoreDesc&listSize=72"
                            self.open(url)
                            

                    

                    query = """
                                INSERT OR IGNORE INTO rankings 
                                    ( date, keyword, page, deltype, rank, prdcode, optcode, dispcode,  dispname,  price, thumb, link, rating, rating_cnt, mine) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                    cur.executemany(query,rankings )
                    con.commit()
                    update_date(keyword)
                    success_cnt += 1

                    
            if success_cnt > 0:
                show_custom_notification_(data_1 = f"{success_cnt} 개의 검색어 랭킹 수집 완료", data_2 = "마진가이드에서 확인하세요.")
        # till 타임 지우기
        delete_endtime()
        return True  
    except:
        delete_endtime()
        show_custom_notification_(data_1 = f"키워드 랭킹검색이 중단되었습니다.", data_2 = "다시 시도해 주세요.")
        return False


if __name__ == "__main__":
    if rankings():
        rankings()
    