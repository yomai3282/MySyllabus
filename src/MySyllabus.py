#MySyllabus.py
import requests,bs4,re,unicodedata,mojimoji,datetime
import pandas as pd
l2s=lambda s,l:s.join(map(str,l))
#strとdatetimeを相互に変換する関数の定義
strptime=lambda str,format:datetime.datetime.strptime(str,format)
strftime=lambda dt,format:datetime.datetime.strftime(dt,format)

date2weekday={0:"月",1:"火",2:"水",3:"木",4:"金",5:"土",6:"日"}
weekday2date={'月':0,'火':1,'水':2,'木':3,'金':4,'土':5,'日':6}

#学期・時限を日付・時刻データに対応させる
def get_sem():
    """
    semester.csvから学期と日付の対応(開始日・終了日)を取得します。
    """
    df = pd.read_csv('../config/semester.csv',index_col="学期")
    sem2date=dict(zip(list(df.index),df.values.tolist()))
    return sem2date

def get_period():
    """
    period.csvから時限と時刻の対応(開始・終了)を取得します。
    """
    df = pd.read_csv('../config/period.csv',index_col="時限")
    period2time=dict(zip(list(df.index),df.values.tolist()))
    return period2time

period2time=get_period() #時限と時刻の対応
sem2date=get_sem() #学期と日付の対応
year=strptime(sem2date[1][0],'%Y/%m/%d').year #今年

# %%
class Course:
    """
    講義情報を扱うクラス
    """
    def __init__(self):
        self.full_id="" #講義番号
        self.title="" #講義名
        self.classroomlist=[] #教室リスト
        self.semlist=[] #開講学期
        self.datelist=[] #曜日・時限
        self.faculty="" #開講学部
        self.department="" #対象学科
        self.description="" #内容
        self.infocols=["講義番号","開講学部","対象学科","講義名","教室","開講学期","曜日時限"]
        self.infovals=[] #講義情報をまとめたリスト
        self.syllabus_url="" #シラバスURL
        self.moodle_url="" #moodleURL
    
    def get_info(self):
        """
        シラバスから情報を取得し、オブジェクトにセットします。
        """
        syllabus_url="https://kyomu.adm.okayama-u.ac.jp/Portal/Public/Syllabus/SyllabusSearchStart.aspx?lct_year="+str(year)+"&lct_cd="+self.full_id+"&je_cd=1"
        moodle_url="https://moodle.el.okayama-u.ac.jp/course/view.php?idnumber="+self.full_id
        res=requests.get(syllabus_url,timeout=1)
        res.raise_for_status() #ウェブページを取得出来ているか確認
        soup=bs4.BeautifulSoup(res.text,"html.parser")
        elems=soup.select("span")
        #テキスト処理
        text=[]
        for elem in elems:
            word=elem.getText().split(',')
            text.append(word)
        #取得した情報をset
        self.faculty=text[7][0]
        if text[8][0]!="昼間":
            self.department=text[8][0]
        self.title=text[11][0] 
        self.semlist=[int(unicodedata.normalize("NFKC",x)) for x in re.findall(r"\d",text[6][0])]
        self.datelist=text[30]
        self.classroomlist=text[22]
        self.syllabus_url=syllabus_url
        self.moodle_url=moodle_url
        self.description+="シラバスURL: "+syllabus_url+"\n"+"moodleURL: "+moodle_url
        self.infovals=[
            self.full_id,
            self.faculty,
            self.department,
            self.title,l2s(',',self.classroomlist),
            l2s(', ',[mojimoji.han_to_zen(str(s)) for s in self.semlist])+"学期",
            l2s(', ',self.datelist)
            ]
    
    def show_info(self):
        """
        オブジェクトにセットされている情報を表示します。
        例：
        講義番号: 2024098456
        講義名:   ＵＮＩＸプログラミング
        教室:     工学部１号館大講義室
        開講学期: １学期
        曜日時限: 月5〜6,木1〜2
        """
        infostr=""
        for col,val in zip(self.infocols,self.infovals):
            infostr+="{:　<5}".format(col)+":　"+val+"\n"
        print(infostr)

# %%
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
import os
import pandas as pd
from dotenv import load_dotenv
import datetime
load_dotenv("../env/.env")

#google calendar_apiのための認証情報
creds = service_account.Credentials.from_service_account_file('../env/credentials.json')
service = build('calendar', 'v3', credentials=creds)
calendar_id = os.getenv('MAIL_ADDRESS')  #自分のメールアドレス


class Schedule(Course):
    """
    グーグルカレンダーの予定を扱うクラスです。
    講義情報を扱うクラス(Course)の派生クラスです。
    """
    def __init__(self):
        super().__init__()
        self.starttime_dict={} #講義の開始時刻
        self.endtime_dict={} #講義の終了時刻
    def get_time(self):
        """
        講義情報から学期・曜日ごとの予定作成時刻を求めます。
        """
        for sem in self.semlist:
            #開講学期でネストする
            self.starttime_dict[sem]={}
            self.endtime_dict[sem]={}
            date_sem=[strptime(sem2date[sem][tf], "%Y/%m/%d") for tf in [0,1]] #第sem学期の日付の始終
            for date in self.datelist:
                #開講曜日でネストする
                self.starttime_dict[sem][date]=[] #第sem学期のdate曜日における開講時刻
                self.endtime_dict[sem][date]=[] #第sem学期のdate曜日における閉講時刻
                try:
                    weekday=re.search(r'[月火水木金]',date).group() #開講曜日
                except:
                    continue
                period=[] #開講時限の始終
                for i in re.findall(r"\d+", date):
                    if len(re.findall(r"\d+", date))>2:
                        if i!=0:
                            period.append(int(i))
                    else:
                        period.append(int(i))
                class_period=[strptime(period2time[period[tf]][tf],"%H:%M") for tf in [0,1]] #開講時刻の始終
                today=datetime.datetime.now() #今日の日付
                #今日の日付を基準に、開講する日付までの曜日のずれを数え、その分日付を進める。(曜日を揃える)
                delta_day=datetime.timedelta(days=(weekday2date[weekday]-today.weekday())%7)
                hour=[class_period[0].hour,class_period[1].hour]
                minute=[class_period[0].minute,class_period[1].minute]
                dt=lambda tf:datetime.datetime(today.year,today.month,today.day,hour[tf],minute[tf])+delta_day
                #曜日を揃えた日付をstartとする。
                start=dt(0)
                end=dt(1)
                while True:
                    #start < 開講学期の初日 => start + 7day
                    if(start<=date_sem[0]):
                        start+=datetime.timedelta(days=7)
                        end+=datetime.timedelta(days=7)
                        continue
                    #開講学期の初日 < start < 開講学期の最終日 => 時刻を格納し、start + 7day
                    elif((date_sem[0]<=start)&(start<=date_sem[1]+datetime.timedelta(days=1))):
                        self.starttime_dict[sem][date].append(start)
                        self.endtime_dict[sem][date].append(end)
                        start+=datetime.timedelta(days=7)
                        end+=datetime.timedelta(days=7)
                        continue 
                    #それ以外ならループを抜ける
                    else:
                        break
    
    def create_event(self,start_time,end_time):
        """
        Googleカレンダーに予定を作成します。
        """
        event = {
            'summary': self.title,
            'location':l2s(',',self.classroomlist),
            'description':self.description,
            'start': {
                'dateTime': strftime(start_time,'%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'Asia/Tokyo',
            },
            'end': {
                'dateTime': strftime(end_time,'%Y-%m-%dT%H:%M:%S'),
                'timeZone': 'Asia/Tokyo',
            },
        }
        event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f'Event created: {event.get("htmlLink")}')

# %%
# サンプルコード
import tkinter as tk
from tkinter import ttk
import pandas as pd
import webbrowser
import time

class Application(tk.Frame):
    """
    ウインドウのクラスです。
    """
    def __init__(self, master=None):
        # Windowの初期設定
        super().__init__(master)
        self.master.title("MySyllabus")
        self.window_width=300
        self.window_height=100
        self.frame_width=870
        self.frame_height=260
        self.master.geometry(f"{self.window_width}x{self.window_height}")
        
        #講義番号入力フォーム
        id_frame=tk.Frame(root,relief="ridge")
        
        #年度
        year_frame=ttk.Labelframe(id_frame,text="講義番号")
        year_frame.pack(side=tk.LEFT)
        
        label1=tk.Label(year_frame,text=f"{year}",width=8)
        label1.pack()
        #中2桁
        id_two_frame=ttk.Labelframe(id_frame,text="中2桁")
        id_two_frame.pack(side=tk.LEFT)
        self.entry1=tk.Entry(id_two_frame,width=4)
        self.entry1.pack()
        #下4桁
        id_six_frame=ttk.Labelframe(id_frame,text="下4桁")
        id_six_frame.pack(side=tk.LEFT)
        self.entry2=tk.Entry(id_six_frame,width=8)
        self.entry2.pack()
        
        #検索ボタン
        id_button_frame=ttk.Labelframe(id_frame)
        id_button_frame.pack(side=tk.LEFT)
        button = tk.Button(id_button_frame,text="検索",command=self.search_info)
        button.pack()
        
        id_frame.pack()
        
        #学務URL
        kyomu_url=tk.Label(root,text="学務情報システムへ",fg="blue")
        kyomu_url.pack()
        kyomu_url.bind("<Button-1>",lambda e:self.link_click("https://kyomu.adm.okayama-u.ac.jp/Portal/StudentApp/Top.aspx"))
        
    def search_info(self):
        """
        講義番号から講義情報を検索します
        """
        try:
            course=Schedule()
            #講義番号
            course.full_id=str(year)+self.entry1.get()+self.entry2.get()
            course.get_info()
            #ターミナルに表示
            course.show_info()
            #GUIに表示
            self.show_info(course)
            #ウインドウサイズを変更
            self.resize_window()
            try:
                self.error.destroy()
                self.add_msg.destroy()
            except:
                pass
        except Exception as e:
            self.error=tk.Frame(root)
            self.error.pack()
            error_msg=tk.Label(self.error,text="講義情報が取得できませんでした")
            error_msg.pack()
            error_detail=tk.Label(self.error,text=f"{e}")
            error_detail.pack()
            
    def make_schedule(self,course):
        """"
        Googleカレンダーに予定を追加します
        """
        for sem in course.semlist:
            for date in course.datelist:
                for x,y in zip(course.starttime_dict[sem][date],course.endtime_dict[sem][date]):
                    course.create_event(x,y)
        self.add_msg=tk.Label(root,text="予定が追加されました!")
        self.add_msg.pack()
        
    def show_info(self,course):
        """
        取得した講義情報を表示します
        """
        #講義情報のLabelframe
        info_frame=tk.Frame(root,relief="ridge")
        course_frame = ttk.Labelframe(
            info_frame,
            relief="ridge",
            text="講義情報",  #タイトルの設定
            labelanchor="n",    #タイトル位置の設定
            )
        
        #label1
        for col,val in zip(course.infocols,course.infovals):
            l = tk.Label(course_frame,text="{:　<5}".format(col)+":　"+val)
            #ラベルを左寄せ
            l.pack(anchor=tk.W)
            
        #講義情報->シラバス・moodleURLのLabelframe
        url_frame=tk.Frame(course_frame)
        #label1
        syllabus_url=tk.Label(url_frame,text="シラバス",fg="blue")
        syllabus_url.pack(side=tk.LEFT)
        syllabus_url.bind("<Button-1>",lambda e:self.link_click(course.syllabus_url))
        #label2
        moodle_url=tk.Label(url_frame,text="moodle",fg="blue")
        moodle_url.pack(side=tk.LEFT)
        moodle_url.bind("<Button-1>",lambda e:self.link_click(course.moodle_url))
        url_frame.pack()
        #space
        space=tk.Label(course_frame,text="")
        space.pack()
        #登録ボタン
        button = tk.Button(course_frame,text="Googleカレンダーに予定を追加",command=lambda:self.make_schedule(course))
        button.pack()
        course_frame.pack(padx=10,pady=10)
        
        button2 = tk.Button(course_frame,text="講義情報を非表示",command=lambda:self.clear_info(info_frame))
        button2.pack()
        course_frame.pack(padx=10,pady=10,side=tk.LEFT)
        
        #開講時刻のLabelframe
        period_frame = ttk.Labelframe(info_frame,relief="ridge",text="開講時刻")
        #時刻を取得
        course.get_time()
        #予定を追加する時刻をラベルとして表示
        for sem in course.semlist:
            for date in course.datelist:
                #開講時刻->曜日毎のLabelframe
                date_frame=ttk.Labelframe(
                    period_frame,
                    relief="ridge",
                    text="第{0}学期 {1}".format(sem,date),
                    labelanchor="n"
                )
                for x,y in zip(course.starttime_dict[sem][date],course.endtime_dict[sem][date]):    
                    s=strftime(x,"%m/%d")+"({})".format(date2weekday[x.weekday()])+" "+strftime(x,"%H:%M")+" ~ "+strftime(y,"%H:%M")
                    #label1
                    time=tk.Label(date_frame,text=s)
                    time.pack()
                date_frame.pack(padx=5,pady=10,side=tk.LEFT)
                
        period_frame.pack(padx=10,pady=10,side=tk.LEFT)
        info_frame.pack()
        
    def clear_info(self,info_frame):
        info_frame.destroy()
        self.window_height-=self.frame_height
    
    def link_click(self,url):
        """
        ハイパーリンクをウェブブラウザで開きます。
        """
        webbrowser.open_new(url)

    #ウインドウサイズを変更
    def resize_window(self):
        self.window_width=self.frame_width
        self.window_height+=self.frame_height
        self.master.geometry(f"{self.window_width}x{self.window_height}")

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()


