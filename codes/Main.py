# coding=utf-8
# 编译日期：2023-12-11 01:37:14
# 版权所有：www.i-search.com.cn
import ubpa.init_input as iinput
from ubpa.base_util import StdOutHook, ExceptionHandler
import ubpa.itools.rpa_str as rpa_str
import ubpa.iexcel as iexcel
import ubpa.ichrome_firefox as ichrome_firefox
import ubpa.ikeyboard as ikeyboard
import ubpa.ics as ics
import ubpa.iwin as iwin
import ubpa.ibrowse as ibrowse
import time
import pdb
from ubpa.ilog import ILog
import getopt
from sys import argv
import sys
import os
from ubpa.base_img import *

class RPA_Project:
     
    def __init__(self,**kwargs):
        self.__logger = ILog(__file__)
        self.path = set_img_res_path(__file__)
        self.robot_no = ''
        self.proc_no = ''
        self.job_no = ''
        self.input_arg = ''
        if('robot_no' in kwargs.keys()):
            self.robot_no = kwargs['robot_no']
        if('proc_no' in kwargs.keys()):
            self.proc_no = kwargs['proc_no']
        if('job_no' in kwargs.keys()):
            self.job_no = kwargs['job_no']
        ILog.JOB_NO, ILog.OLD_STDOUT = self.job_no, sys.stdout
        sys.stdout = StdOutHook(self.job_no, sys.stdout)
        ExceptionHandler.JOB_NO, ExceptionHandler.OLD_STDERR = self.job_no, sys.stderr
        sys.excepthook = ExceptionHandler.handle_exception
        if('input_arg' in kwargs.keys()):
            self.input_arg = kwargs['input_arg']
            if(len(self.input_arg) <= 0):
                self.input_arg = iinput.load_init(__file__)
            if self.input_arg is None:
                sys.exit(0)
      
    def Main(self,pv_1=None):
        #打开浏览器
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2023121023491311564,Title:打开浏览器,Note:')
        tvar_2023121023491311565=ibrowse.open_browser(browser_type='chrome',url='https://cmbchina.com/cfweb/personal/default.aspx',maximum=0)
        print('[Main] [打开浏览器] [SNTag:2023121023491311564]  返回值：[' + str(type(tvar_2023121023491311565)) + ']' + str(tvar_2023121023491311565))
        #最大化窗口
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2023121023500550787,Title:最大化窗口,Note:')
        tvar_2023121023500550788=iwin.do_win_maximize(waitfor=10.000,win_title=r'招商银行 -- 个人理财产品 - Google Chrome',win_class=r'Chrome_WidgetWin_1')
        print('[Main] [最大化窗口] [SNTag:2023121023500550787]  返回值：[' + str(type(tvar_2023121023500550788)) + ']' + str(tvar_2023121023500550788))
        #结构化抓取
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231210235147865207,Title:结构化抓取,Note:')
        tvar_20231210235530677210=ichrome_firefox.extract_data_chrome(title=r'招商银行 -- 个人理财产品',url=r'https://cmbchina.com/*',metaData=[{"attr":["text"],"exact":1,"name":"Column0","path":[{"tag":"div"}]}],attrMap={"hasReachedRelativeAncestor":"false","isTable":"false","nodeHierarchyInfo":[{"isPresentInSelector":1,"otherAttributes":{},"selectorInfo":{"attributes":{"css-selector":"#cList","id":"cList","tag":"DIV","xpath":"//*[@id=\"cList\"]"},"index":0,"tagName":"DIV"}}]},run_mode='unctrl',columns_setting=[{"DataType":"text","IsTable":0,"IsVisible":1,"Name":"Column0","ReferenceName":"Column0","Value":"中银稳富封闭23147期-387天（和盈）\n代码：ZY010267\n产品到期日：2025-01-09（期限：387天）\n业绩比较基准区间：3.10%至3.60%\n风险评级：中低风险\n发售渠道：网银专业版|网银大众版|手机|全球连线手机|全球连线网银\n发售地：全行\n理财产品登记编码：Z7001023000531\n发售起始日：2023-12-12\n发售截止日：2023-12-18\n业绩比较基准不代表未来表现和实际收益。本理财产品为固定收益类产品，以产品投资货币市场工具仓位0%-20%，信用债仓位0%-80%，非标准化债权类资产仓位0%-50%，组合杠杆率115%为例，业绩比较基准参考本产品发行时已知的中债-综合财富（1-3年）指数收益率、期限匹配的非标准化债权类资产收益率，考虑本理财产品综合费率、资本利得收益并结合产品投资策略进行测算得出。\n查看更多 >","attr":"text"}],extract_mode=0,max_limit=500,img_res_path = self.path)
        print('[Main] [结构化抓取] [SNTag:20231210235147865207]  返回值：[' + str(type(tvar_20231210235530677210)) + ']' + str(tvar_20231210235530677210))
        #单元格写入
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231210235709619273,Title:单元格写入,Note:')
        tvar_20231210235709619274=iexcel.write_cell(path='C:/Users/zhguo/Desktop/营销产品.xlsx',text=tvar_20231210235530677210,cell='K20')
        print('[Main] [单元格写入] [SNTag:20231210235709619273]  返回值：[' + str(type(tvar_20231210235709619274)) + ']' + str(tvar_20231210235709619274))
        #单元格读取
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2020053011133768956,Title:单元格读取,Note:')
        tvar_2020053011133769057=iexcel.read_cell(path='C:/Users/zhguo/Desktop/营销产品.xlsx',cell='K2',cell_type=None)
        print('[Main] [单元格读取] [SNTag:2020053011133768956]  返回值：[' + str(type(tvar_2020053011133769057)) + ']' + str(tvar_2020053011133769057))
        #打开浏览器
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231105190332863238,Title:打开浏览器,Note:')
        tvar_20231105190332863239=ibrowse.open_browser(browser_type='chrome',url='https://yiyan.baidu.com/',maximum=0)
        print('[Main] [打开浏览器] [SNTag:20231105190332863238]  返回值：[' + str(type(tvar_20231105190332863239)) + ']' + str(tvar_20231105190332863239))
        #最大化窗口
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231105190955298371,Title:最大化窗口,Note:')
        tvar_20231105190955298372=iwin.do_win_maximize(waitfor=10.000,win_title=r'文心一言 - Google Chrome',win_class=r'Chrome_WidgetWin_1')
        print('[Main] [最大化窗口] [SNTag:20231105190955298371]  返回值：[' + str(type(tvar_20231105190955298372)) + ']' + str(tvar_20231105190955298372))
        #点击
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231105190536381312,Title:点击,Note:')
        ics.truple_mouse_click(distpos=(1120,1375),win_title=r'文心一言 - Google Chrome',win_class=r'Chrome_WidgetWin_1')
        #热键输入
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231105192046619464,Title:热键输入,Note:')
        ikeyboard.key_send_cs(waitfor=10.000,text=tvar_2020053011133769057)
        time.sleep(40)
        #结构化抓取
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:20231105191452842417,Title:结构化抓取,Note:')
        tvar_20231105191537008421=ichrome_firefox.extract_data_chrome(title=r'文心一言',url=r'https://yiyan.baidu.com/*',metaData=[{"attr":["text"],"exact":1,"name":"Column0","path":[{"tag":"div"}]}],attrMap={"hasReachedRelativeAncestor":"false","isTable":"false","nodeHierarchyInfo":[{"isPresentInSelector":1,"otherAttributes":{},"selectorInfo":{"attributes":{"css-selector":"#DIALOGUE_CARD_LIST_ID > div:nth-child(1) > div:nth-child(3) > div:nth-child(1) > div:nth-child(2)","parentid":"DIALOGUE_CARD_LIST_ID","tag":"DIV","xpath":"//*[@id=\"DIALOGUE_CARD_LIST_ID\"]/div[1]/div[3]/div[1]/div[2]"},"index":0,"tagName":"DIV"}}],"windowname":"__bid_n=186cb2d2ef16ffcd984207"},run_mode='unctrl',columns_setting=[{"DataType":"text","IsTable":0,"IsVisible":1,"Name":"Column0","ReferenceName":"Column0","Value":"Spring has arrived, bringing new life to the world around us. Just as nature blooms and flourishes, your finances have the opportunity to grow and thrive. Introducing the \"Bank of China Steady Wealth Closed-End 23147-387 Days (Heying)\" product, code ZY010267.\n\nWith medium-low risk, this fixed-income financial product offers a secure opportunity for growth. With a product number of Z7001023000531, a mature date of January 9th, 2025, and a duration of 387 days, the investment period begins on December 12th, 2023, and ends on December 18th, 2023.\n\nThe performance comparison base range is between 3.10% and 3.60%, offering competitive returns on your investment. However, it?s important to note that the performance comparison base does not represent future performance or actual returns.\n\nThis financial product invests in a mix of monetary market instruments (0%-20%), credit bonds (0%-80%), and non-standardized debt assets (0%-50%), with a portfolio leverage ratio of 115%. The performance comparison base is calculated based on the known yield of the China Bond - Composite Wealth (1-3 years) Index at the time of issuance, the yield of non-standardized debt assets with matching maturities, and considering the comprehensive fee rate, capital gains income, and product investment strategy of this financial product.\n\nTake advantage of the season of growth and renewal by investing in the Bank of China Steady Wealth Closed-End 23147-387 Days (Heying) product today. Spring is a time for new beginnings, and this financial product offers a fresh start for your financial journey. Don?t miss out on the opportunity to cultivate wealth and secure your financial future. Act now and let your money grow with the warmth and vitality of spring.","attr":"text"}],extract_mode=0,max_limit=500,img_res_path = self.path)
        print('[Main] [结构化抓取] [SNTag:20231105191452842417]  返回值：[' + str(type(tvar_20231105191537008421)) + ']' + str(tvar_20231105191537008421))
        #单元格写入
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2020053011193201962,Title:单元格写入,Note:')
        tvar_2020053011193201963=iexcel.write_cell(path='C:/Users/zhguo/Desktop/营销产品.xlsx',text='Spring has arrived, bringing new life to the world around us. Just as nature rejuvenates, its time to rejuvenate your finances with Bank of Chinas stable and prosperous wealth-building product—the Zhongyin Wenfu Fengbi 23147 Period - 387 Days (Heying) with code ZY010267.Embrace the season of growth and let your investments blossom with this medium-low risk product, tailored to deliver fixed income. With a product number Z7001023000531, its designed to mature on 2025-01-09, giving you the perfect 387-day investment horizon. The subscription period begins on 2023/12/12 and ends on 2023/12/18, providing you with a narrow window of opportunity to seize the moment.Anticipate a performance comparison base range of 3.10% to 3.60%. However, its essential to understand that this benchmark doesnt guarantee future performance or actual returns. Our financial product primarily invests in monetary market instruments (0%-20%), credit bonds (0%-80%), and non-standardized debt assets (0%-50%), with a portfolio leverage ratio of 115%.The performance comparison benchmark is derived from a meticulous calculation, considering the known yield of the China Bond - Composite Wealth (1-3 years) Index at the time of the products issuance, the yield of non-standardized debt assets matching the investment period, the comprehensive fee rate of the financial product, capital gains, and the product investment strategy.Harness the power of spring and let your finances flourish with the Bank of China Zhongyin Wenfu Fengbi 23147 Period - 387 Days (Heying). Invest wisely, invest in harmony with the seasons, and watch your wealth grow like the blooming flowers of spring. Seize this opportunity before its too late!Disclaimer: The performance comparison benchmark does not represent future performance or actual returns. This financial product is a fixed-income instrument, and investors should carefully study and evaluate the investment risks before making an investment decision..',cell='M2')
        print('[Main] [单元格写入] [SNTag:2020053011193201962]  返回值：[' + str(type(tvar_2020053011193201963)) + ']' + str(tvar_2020053011193201963))
        #输出
        self.__logger.dlogs(job_no=self.job_no,logmsg='Flow:Main,StepNodeTag:2020053011200266865,Title:输出,Note:')
        rpa_str.iprints('写入成功，快去新打开的表格中查看单元格M2写入的内容吧！')
 
if __name__ == '__main__':
    ILog.begin_init()
    robot_no = ''
    proc_no = ''
    job_no = ''
    input_arg = ''
    try:
        argv = sys.argv[1:]
        opts, args = getopt.getopt(argv,"hr:p:j:i:",["robot = ","proc = ","job = ","input = "])
    except getopt.GetoptError:
        print ('robot.py -r <robot> -p <proc> -j <job>')
    for opt, arg in opts:
        if opt == '-h':
            print ('robot.py -r <robot> -p <proc> -j <job>')
        elif opt in ("-r", "--robot"):
            robot_no = arg
        elif opt in ("-p", "--proc"):
            proc_no = arg
        elif opt in ("-j", "--job"):
            job_no = arg
        elif opt in ("-i", "--input"):
            input_arg = arg
    pro = RPA_Project(robot_no=robot_no,proc_no=proc_no,job_no=job_no,input_arg=input_arg)
    pro.Main()
