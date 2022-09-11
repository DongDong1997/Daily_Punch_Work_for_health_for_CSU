# -*- coding:gbk -*-


import winreg
import wx
import sqlite3
from  platform import architecture
from wx.adv import AnimationCtrl
from selenium import webdriver
from os import popen,remove
import time
import shutil
from random import randint
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from msedge.selenium_tools import EdgeOptions
import datetime


#class WarnMesFont(wx.Font): #主页面提醒语句
#    def __init__(self,parent):
#        super(WarnMesFont,self).__init__(parent)
#        self.PointSize = 15
        
class LoadDialog(wx.Dialog):
    def __init__(self,parent,title):
        super(LoadDialog,self).__init__(parent,title=title,size = (250,150))
        self.con1 = sqlite3.connect('res/INFO.db')
        self.cur1 = self.con1.cursor()
        self.content = self.cur1.execute("SELECT ID,PW FROM INFO").fetchall()
        a1 = self.content[0]
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        hbox1.AddSpacer(20)
        text1 = wx.StaticText(panel,label = '学号:')
        hbox1.Add(text1,0, wx.ALIGN_CENTRE|wx.ALL)
        self.text2 = wx.TextCtrl(panel,1,value = str(a1[0]),size = (160,-1))
        hbox1.Add(self.text2,0, wx.ALIGN_CENTRE|wx.ALL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        text3 = wx.StaticText(panel,label = '密码:')
        hbox2.AddSpacer(20)
        hbox2.Add(text3,0, wx.ALIGN_CENTRE|wx.ALL)
        self.text4 = wx.TextCtrl(panel,2,value = str(a1[1]),size= (160,-1))
        hbox2.Add(self.text4,0, wx.ALIGN_CENTRE|wx.ALL)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.btn1 = wx.Button(panel,label = "保存")
        self.btn1.Bind(wx.EVT_BUTTON,self.Save)
        hbox3.AddSpacer(50)
        hbox3.Add(self.btn1,0,wx.ALIGN_CENTER|wx.ALL)
        self.btn2 = wx.Button(panel,label = "取消")
        self.btn2.Bind(wx.EVT_BUTTON,self.Quit)
        hbox3.Add(self.btn2,0,wx.ALIGN_CENTER|wx.ALL)
        vbox.Add(hbox1,wx.ALIGN_CENTRE)
        vbox.Add(hbox2,wx.ALIGN_CENTRE)
        vbox.Add(hbox3,wx.ALIGN_CENTRE)
        panel.SetSizer(vbox)
        self.SetIcon(wx.Icon('res/Load.ico'))
        self.Centre()
        

    
    def Save(self,event):
        wx.MessageBox("账号保存完毕","提醒",wx.OK|wx.ICON_INFORMATION)
        self.cur1.execute("UPDATE INFO SET ID = ?, PW = ?",(self.text2.GetValue(),self.text4.GetValue()))
        self.con1.commit()
        print('账号保存')
    
    def Quit(self,event):
        self.cur1.close()
        self.con1.close()
        print("账号取消")
        self.Destroy()
        
   # def UseDialog(self):
    #    panel = wx.Panel(Self)
class TimeDialog(wx.Dialog):
    def __init__(self,parent,title):
        super(TimeDialog,self).__init__(parent,title = title,size = (200,150))
        self.con1 = sqlite3.connect('res/INFO.db')
        self.cur1 = self.con1.cursor()
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        self.content = self.cur1.execute("SELECT H,M,S FROM INFO").fetchall()[0]
        #print(self.content)
        text1 = wx.StaticText(panel,label = '时',style = wx.ALIGN_CENTER_HORIZONTAL)
        hbox1.Add(text1,0,wx.ALIGN_CENTER|wx.ALL)
        hbox1.Add((42,0))
        text2 = wx.StaticText(panel,label = '分',style = wx.ALIGN_CENTER_HORIZONTAL)
        hbox1.Add(text2,0,wx.ALIGN_CENTER|wx.ALL)
        hbox1.Add((42,0))
        text3 = wx.StaticText(panel,label = "秒",style = wx.ALIGN_CENTER_HORIZONTAL)
        hbox1.Add(text3,0,wx.ALIGN_CENTER|wx.ALL)
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.text4 = wx.TextCtrl(panel,1,value = str(self.content[0]),size = (50,-1),style = wx.ALIGN_RIGHT)
        self.text5 = wx.TextCtrl(panel,2,value = str(self.content[1]),size = (50,-1),style = wx.ALIGN_RIGHT)
        self.text6 = wx.TextCtrl(panel,3,value = str(self.content[2]),size = (50,-1),style = wx.ALIGN_RIGHT)
        t1 = wx.StaticText(panel,-1,":",style = wx.ALIGN_LEFT)
        t2 = wx.StaticText(panel,-1,":",style = wx.ALIGN_LEFT)
        hbox2.Add(self.text4,0,wx.ALIGN_LEFT)
        hbox2.Add(t1,0,wx.ALIGN_LEFT)
        hbox2.Add(self.text5,0,wx.ALIGN_LEFT)
        hbox2.Add(t2,0,wx.ALIGN_LEFT)
        hbox2.Add(self.text6,0,wx.ALIGN_LEFT)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.btn1 = wx.Button(panel,label = "保存")
        self.btn1.Bind(wx.EVT_BUTTON,self.Save)
        #self.btn1.Bind(wx.EVT_TEXT,self.Save)
        hbox3.Add(self.btn1,0,wx.ALIGN_CENTER|wx.ALL,5)
        self.btn2 = wx.Button(panel,label = "取消")
        self.btn2.Bind(wx.EVT_BUTTON,self.Quit)
        hbox3.Add(self.btn2,0,wx.ALIGN_CENTER|wx.ALL,5)
        vbox.Add(hbox1,0,wx.ALIGN_CENTER)
        vbox.Add(hbox2,1,wx.ALIGN_CENTER)
        vbox.Add(hbox3,2,wx.ALIGN_CENTER)
        panel.SetSizer(vbox)
        self.SetIcon(wx.Icon('res/Time.ico'))
        self.Centre()
        

    
    def Save(self,event):
        a= wx.MessageBox("保存完毕","提醒",wx.OK|wx.ICON_INFORMATION)
        self.cur1.execute("UPDATE INFO SET H = ?,M = ?,S = ?",(self.text4.GetValue(),self.text5.GetValue(),self.text6.GetValue()))
        print(self.text4.GetValue())
        self.con1.commit()
        print('保存完毕')
    
    def Quit(self,event):
        self.cur1.close()
        self.con1.close()
        print("取消")
        self.Destroy()
        
class NoteDialog(wx.Dialog):
    def __init__(self,parent,title):
        super(NoteDialog,self).__init__(parent,title =title, size = (300,200))
        self.con1 = sqlite3.connect("res/TEST.db")
        self.cur1 = self.con1.cursor()
        nb = wx.Listbook(self)
        MyPanel1 = MyPanel(nb)
        MyPanel2 = MyPanel(nb)
        nb.AddPage(MyPanel1,"日志")
        nb.AddPage(MyPanel2,"具体修改")
        content1 = self.cur1.execute("SELECT YEAR,MONTH,DAY,MES,CODE FROM TEST").fetchall()
        text1 = ""
        text2 = ""
        for each in content1:
            text1 = text1 + each[0] + '年' + each[1] + '月' + each[2] + '日' + '\n'+ each[3] + '\n'
            text2 = text2 + each[0] + '年' + each[1] + '月' + each[2] + '日' + '\n'+ each[4] + '\n'
        MyPanel1.SetText(text1)
        MyPanel2.SetText(text2)
        self.cur1.close()
        self.con1.close()
        self.Centre()
        self.SetIcon(wx.Icon('res/Note.ico'))
        self.Show(True)
        
class MyPanel(wx.Panel):
    def __init__(self,parent):
        super(MyPanel,self).__init__(parent)
        self.text = wx.StaticText(self,style = wx.ALIGN_CENTER_HORIZONTAL)
        
        
    def SetText(self,MyText):
        self.text.SetLabel(MyText)
        
        
class UseDialog(wx.Dialog):
    def __init__(self,parent,title):
        super(UseDialog,self).__init__(parent,title = title,size = (400,200))
        self.con1 = sqlite3.connect('res/INFO.db')
        self.cur1 = self.con1.cursor()
        nb = wx.Listbook(self)
        MyPanel1 = MyPanel(nb)
        MyPanel2 = MyPanel(nb)
        MyPanel3 = MyPanel(nb)
        nb.AddPage(MyPanel1,"账号设置")
        nb.AddPage(MyPanel2,"时间设置")
        nb.AddPage(MyPanel3,"注意事项")
        self.content = self.cur1.execute('SELECT ACC,CLOCK,WARN FROM MUSE').fetchall()
        MyPanel1.SetText(str(self.content[0][0]))
        MyPanel2.SetText(str(self.content[0][1]))
        MyPanel3.SetText(str(self.content[0][2]))
        self.con1.commit()
        #self.cur1.close()
        #self.con1.close()
        self.Centre()
        self.SetIcon(wx.Icon('res/Use.ico'))
        self.Show(True)

class UpdateDialog(wx.Dialog):
    def __init__(self,parent,title):
        super(UpdateDialog,self).__init__(parent,title = title ,size = (250,170))
        panel = wx.Panel(self)
        vbox1 = wx.BoxSizer(wx.VERTICAL)
        t1= GetEdgeVer()
        t2= GetLocalDriverVer()
        vbox1.AddSpacer(20)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        localtext1 = wx.StaticText(panel,1,label='浏览器版本:',style = wx.ALIGN_CENTER)
        hbox1.Add(localtext1,0, wx.ALIGN_LEFT|wx.ALL)
        self.localvertext = wx.StaticText(panel,3,style = wx.ALIGN_CENTER)
        hbox1.Add(self.localvertext,0,wx.ALIGN_LEFT)
        vbox1.Add(hbox1,0,wx.ALIGN_CENTER|wx.ALL)
        self.localvertext.SetLabel(t1)
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        drivertext1 = wx.StaticText(panel,2,label=" 驱动版本 :",style = wx.ALIGN_CENTER)
        hbox2.Add(drivertext1,0,wx.ALIGN_LEFT)
        self.driververtext = wx.StaticText(panel,4,style = wx.ALIGN_CENTER)
        hbox2.Add(self.driververtext,0,wx.ALIGN_LEFT)
        self.driververtext.SetLabel(t2)
        vbox1.Add(hbox2,0,wx.ALIGN_CENTER|wx.ALL)
        
        vbox1.AddSpacer(10)
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.btn   = wx.Button(panel,1,size=(150,40))
        hbox3.Add(self.btn, 0,wx.ALIGN_CENTER)
        vbox1.Add(hbox3, 0,wx.ALIGN_CENTER|wx.ALL)
        
        
        self.btn.Bind(wx.EVT_BUTTON,self.OnUpdatever)
        panel.SetSizer(vbox1)
        self.SetIcon(wx.Icon('res/update.ico'))
        self.VerOK = 1 #适配
        self.Centre()
        if abs(int(t1.split('.')[3])-int(t2.split('.')[3])) <= 2:
            self.btn.SetLabel('版本适配')
            self.VerOK = 1
        else:
            self.btn.SetLabel("版本不适配\n[点击更新驱动]")
            self.VerOK = 0
            # 更新适配版本
        self.Show()

    def OnUpdatever(self,event):
        # 更新适配版本
        # edge驱动下载一系类文件夹，提取驱动文件到res，删除其余文件夹
        if self.VerOK == 0:
            self.btn.SetLabel("驱动更新ing")
            EdgeChromiumDriverManager(path = r"res").install()
            Edge_version = GetEdgeVer()
            text = Edge_version.split('.')
            text1= text[0]+'.'+ text[1]+'.'+text[2]  #文件夹只取前三数
            FrameWorkVer= architecture()[0][:2]
            driverpath = 'res\.wdm\drivers\edgedriver\win'+ FrameWorkVer + '\\' + text1
            savepath = r'res'
            remove('res\msedgedriver.exe')
            shutil.move(driverpath+'\msedgedriver.exe',savepath)
            shutil.rmtree('res\.wdm')
            #shutil.rmtree('res\Driver_Notes')
            self.localvertext.SetLabel(GetEdgeVer())
            self.driververtext.SetLabel(GetLocalDriverVer())
            self.btn.SetLabel("版本更新完毕")
            self.VerOK = 1

        
def GetEdgeVer(): #获取电脑Edge的版本号
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'SoftWare\Microsoft\Edge\BLBeacon',0,winreg.KEY_READ)
    Edge_version = winreg.QueryValueEx(key,'version')[0]
    return Edge_version

def GetLocalDriverVer():#获取本地Edge驱动的版本号
    Driver_version = popen(r'res\msedgedriver --V').read().split(' ')[3]
    return Driver_version

# 重写Cron定时
class my_CronTrigger(CronTrigger):
    def __init__(self, year=None, month=None, day=None, week=None, day_of_week=None, hour=None,
                  minute=None, second=None, start_date=None, end_date=None, timezone=None,
                  jitter=None):
         super().__init__(year=None, month=None, day=None, week=None, day_of_week=None, hour=None,
                  minute=None, second=None, start_date=None, end_date=None, timezone=None,
                  jitter=None)
    @classmethod
    def my_from_crontab(cls, expr, timezone=None):
        values = expr.split(' ')
        if len(values) != 7:
            raise ValueError('Wrong number of fields; got {}, expected 7'.format(len(values)))

        return cls(second=values[0], minute=values[1], hour=values[2], day=values[3], month=values[4],
                   day_of_week=values[5], year=values[6], timezone=timezone)
    

#def AutoPunch():
#    driver = webdriver.chrome()   
class MainWin(wx.Frame):
    def __init__(self,parent,title):
        super(MainWin,self).__init__(parent,title = title , size = (400,250))
        self.InitUI()
        
    def InitUI(self):
        # 菜单设置
        Menubar  = wx.MenuBar()
        self.SettingMenu = wx.Menu()
        self.LoadItem = wx.MenuItem(self.SettingMenu,100,text = "登录设置", kind = wx.ITEM_NORMAL)
        self.LoadItem.SetBitmap(wx.Bitmap('res/load.bmp'))
        self.SettingMenu.AppendItem(self.LoadItem)
        self.TimeItem = wx.MenuItem(self.SettingMenu,200,text = "时间设置", kind = wx.ITEM_NORMAL)
        self.TimeItem.SetBitmap(wx.Bitmap('res/time.bmp'))
        self.SettingMenu.AppendItem(self.TimeItem)
        self.SettingMenu.AppendSeparator()
        self.NoteItem = wx.MenuItem(self.SettingMenu,300,text = "更新日志", kind = wx.ITEM_NORMAL)
        self.NoteItem.SetBitmap(wx.Bitmap('res/note.bmp'))
        self.SettingMenu.AppendItem(self.NoteItem)
        self.UseItem  = wx.MenuItem(self.SettingMenu,400,text = "使用说明", kind = wx.ITEM_NORMAL)
        self.UseItem.SetBitmap(wx.Bitmap('res/Use.bmp'))
        self.SettingMenu.AppendItem(self.UseItem)
        self.UpdateItem  = wx.MenuItem(self.SettingMenu,500,text = "驱动更新", kind = wx.ITEM_NORMAL)
        self.UpdateItem.SetBitmap(wx.Bitmap('res/update.bmp'))
        self.SettingMenu.AppendItem(self.UpdateItem)
       
        Menubar.Append(self.SettingMenu,"设置")
        self.SetMenuBar(Menubar)
        self.Bind(wx.EVT_MENU,self.OnLoad,id = 100)
        self.Bind(wx.EVT_MENU,self.OnTime,id = 200)
        self.Bind(wx.EVT_MENU,self.OnNote,id = 300)
        self.Bind(wx.EVT_MENU,self.OnUse,id = 400)
        self.Bind(wx.EVT_MENU,self.OnUpdate,id= 500)
        
        # 主界面布局
        panel = wx.Panel(self)
        VBox1 = wx.BoxSizer(wx.VERTICAL)
        HBox1 = wx.BoxSizer(wx.HORIZONTAL)
        self.ani = AnimationCtrl(panel,1)
        self.ani.LoadFile(r"res/sleep1.gif")
        #HBox1.Add(self.ani,wx.ID_ANY,flag = wx.ALIGN_CENTER)
        #VBox1.Add(HBox1,wx.ALIGN_CENTER)
        VBox1.AddSpacer(10)
        VBox1.Add(self.ani,flag = wx.ALIGN_CENTRE)
        VBox1.AddSpacer(10)
        HBox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.Bt1 = wx.Button(panel,1,'启动')
        self.Bt2 = wx.Button(panel,2,'停止')
        HBox2.Add(self.Bt1,wx.ALIGN_LEFT|wx.ALL)
        HBox2.AddSpacer(30)
        HBox2.Add(self.Bt2,wx.ALIGN_LEFT|wx.ALL)
        VBox1.Add(HBox2,flag = wx.ALIGN_CENTRE)
        VBox1.AddSpacer(10)
        HBox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.WarnMes = wx.StaticText(panel,1,style= wx.ALIGN_CENTER,size = (200,40))
        self.WarnMes.SetLabel("打卡任务已停止")
        
        #self.WarnMes.SetFont(wx.Font(pointsize = 10))
        HBox3.Add(self.WarnMes,3,wx.ALL)
        VBox1.Add(HBox3,flag = wx.ALIGN_CENTRE)
        HBox4 = wx.BoxSizer(wx.HORIZONTAL)
        mytext = wx.StaticText(panel,-1,label = 'by 永远吃不饱的胖达',style = wx.ALL)
        HBox4.Add(mytext,0,wx.RIGHT,2)
        VBox1.Add(HBox4)
        
        self.Bt1.Bind(wx.EVT_BUTTON,self.OnPunch)
        self.Bt2.Bind(wx.EVT_BUTTON,self.OnPause)
        panel.SetSizer(VBox1)
        self.ani.Play()
        
        self.SetSize((300,300))
        self.Centre()
        self.SetIcon(wx.Icon('res/clock.ico'))
        self.Show(True)
        
    def OnLoad(self,event):
        a = LoadDialog(self,"登录设置").ShowModal()
        
    def OnTime(self,event):
        a = TimeDialog(self,"时间设置").ShowModal()
        
    def OnNote(self,event):
        a = NoteDialog(self,"更新日志").Show()
        
    def OnUse(self,event):
        a = UseDialog(self,"使用说明").Show()
        
    def OnUpdate(self,event):
        a = UpdateDialog(self,"驱动更新").ShowModal()
        
    def OnPunch(self,event):
        # 定时任务
        self.WarnMes.SetLabel('任务进行中')
        self.con = sqlite3.connect('res/INFO.db')
        self.cur = self.con.cursor()
        TimeInfo= self.cur.execute("SELECT H,M,S FROM INFO").fetchall()[0]
        self.scheduler = BackgroundScheduler()
        #cronStart = str(TimeInfo[2]) + ' ' + str(TimeInfo[1]) + ' ' + str(TimeInfo[0]) + ' * * * *'
        #print(cronStart)
        self.con.commit()
        self.scheduler.add_job(self.Job,'cron',day_of_week = '*',hour = int(TimeInfo[0]) ,minute = int(TimeInfo[1]))
        self.scheduler.add_job(self.JobUpdate,'cron',day_of_week = '*',hour = 18 ,minute = 51)
        #con.close()
        #cur.close()
        #self.scheduler.add_job(self.NotePunch(),my_CronTrigger.my_from_crontab(cronStart))
        self.scheduler.start()
        
    def OnPause(self,event):
        # 暂停定时任务
        print('任务暂停') 
        self.scheduler.shutdown(wait=False)
        self.WarnMes.SetLabel('打卡任务已停止')
        #self.ani.Stop() 
        #self.con.close()
        #self.cur.close()  
        
    def Job(self):
        self.NotePunch()
        
    def JobUpdate(self):
        self.WarnMes.SetLabel('今日还未打卡')


    def NotePunch(self):
        # 首先是判断驱动是否适配
        #self.ani.Play()
        tok = GetEdgeVer()
        tic = GetLocalDriverVer()
        self.WarnMes.SetLabel('驱动版本检查中')
        time.sleep(1)
        if abs(int(tok.split('.')[3])-int(tic.split('.')[3]))>2:
            self.WarnMes.SetLabel('驱动不适配，请手动更新驱动\n[设置-驱动更新]')
            return
        else:
            self.WarnMes.SetLabel('驱动适配')
        time.sleep(1)
        self.WarnMes.SetLabel('开始执行任务')
        time.sleep(1)
        self.WarnMes.SetLabel('打卡ing')
        url = "https://wxxy.csu.edu.cn/ncov/wap/default/index"
        self.sign_in(url)
        self.WarnMes.SetLabel('今日打卡任务完成')
          
    
    def sign_in(self,url):
        
        
        #个人账号信息导出
        
        conn = sqlite3.connect('res\INFO.db')
        cur = conn.cursor()
        
        idContent = cur.execute("SELECT ID,PW FROM INFO").fetchall()[0]

        
    # conn.close()
    # cur.close()
        # 浏览器设置
        '''
        options = {"ms:edgeOptions":{
                "args":["--disable-gpu-program-cache",
                        "--disable-gpu",
                        "--disable-javascript",
                        "--headless"
            ]
        }
         }
        '''
        '''
        prefs = {
            'profile.default_content_setting_values':{
                'images':2,    #不加载图片
                'javascrip':2  #不进行javascrip渲染
            }
        }
        options.add_experimental_option('prefs',prefs)
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        options.add_argument('--hide-scrollbars')
        options.add_argument('--headless')
        '''
        
        caps= {
                    "browserName": "MicrosoftEdge",
                    "version": "",
                    "platform": "WINDOWS",
                    # 关键是下面这个
                    "ms:edgeOptions": {
                        'extensions': [],
                        'args': [
                            #'--headless',
                            '--disable-gpu',
                            '--remote-debugging-port=9222',
                        ]}
                }
        
        option = EdgeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        driver = webdriver.Edge(r'.\res\msedgedriver.exe',capabilities=caps)
        

        driver.get(url)

        now_handle = driver.current_window_handle
        driver.switch_to.window(now_handle)
        #windows = driver.window_handles
        #driver.switch_to.window(windows[-1])
        driver.find_element_by_id("username").clear()
        driver.find_element_by_id("password").clear()
        self.WarnMes.SetLabel('填写学号、密码')
        driver.find_element_by_id("username").send_keys(str(idContent[0]))
        driver.find_element_by_id("password").send_keys(str(idContent[1]))
        time.sleep(randint(1,2))
        self.WarnMes.SetLabel('登录ing')
        driver.find_element_by_id("login_submit").click()
        time.sleep(randint(4,5))

        now_handle = driver.current_window_handle
        driver.switch_to.window(now_handle)

        try:
            self.WarnMes.SetLabel('登录成功')
            time.sleep(randint(1,2))
            self.WarnMes.SetLabel('点击位置获取')
            time.sleep(1)
            self.WarnMes.SetLabel('获取位置ing')
            driver.find_element_by_xpath('/html/body/div[1]/div/div/section/div[4]/ul/li[14]/div/input').click()
            time.sleep(8)
            try:
                self.WarnMes.SetLabel('提交信息')
                driver.find_element_by_xpath('/html/body/div[1]/div/div/section/div[5]/div/a').click()
                #/html/body/div[1]/div/div/section/div[5]/div/a
                time.sleep(randint(2,3))
                self.WarnMes.SetLabel('确认信息')
                driver.find_element_by_xpath('//*[@id="wapcf"]/div/div[2]/div[2]').click()
                #//*[@id="wapcf"]/div/div[2]/div[2]
                time.sleep(randint(2,3))
                self.WarnMes.SetLabel('打卡完成')
                self.DebugText = '打卡成功'
            except BaseException:
                driver.find_element_by_id('//*[@id="wapat"]/div/div[2]/div').click()
                self.WarnMes.SetLabel('已经打卡过了')
                time.sleep(3)
                self.DebugText = '已经打卡过了'
        except BaseException:
            self.WarnMes.SetLabel('打卡失败')
            self.DebugText = '打卡失败'
        self.DebugText = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ':  ' +  self.DebugText + '\n'
        Debugfile = open('res/Debug.txt',mode='a')
        Debugfile.write(self.DebugText)
        Debugfile.close()
        driver.quit()
    
        
            
if __name__ ==  "__main__":
    ex = wx.App()
    MainWin(None,'中南大学自动打卡')
    ex.MainLoop()
        
        