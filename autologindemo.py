# -*- coding: utf-8 -*-
from splinter.browser import Browser
from PIL  import Image
import pytesseract
import time
import xlrd

url='登录网址'
#登录后跳转网址
dlurl=''
#账号
user=''
#密码
pwd=''
img='jt.png'
excPath=r'D:\Python Code\test.xlsx'
browser = Browser('chrome')
dic={}
keys=[]
zhi=[]
n=0
def  login(): 
	
	browser.visit(url)
	#截图保存
	browser.driver.save_screenshot(img)
	yzmp=Image.open(img)
	box=(645,290,745,328)
	yzm=yzmp.crop(box)
	#提取图片数字验证码
	mstr=pytesseract.image_to_string(yzm)
	
	#填充用户名、密码、验证码
	browser.fill("loginId",user)
	browser.fill("clear_password",pwd)
	browser.fill("loginYzCode",mstr)
	
	browser.find_by_value("登   录").click()
	
#模拟自动输入函数	
def dlinput(jgName,code):

	#根据输入框的name找到并填充搜索框
	browser.fill("search_LIKE_fullname",jgName)
	#根据div的value值找到div
	browser.find_by_value(u"搜索").click() 
	
	
	
	time.sleep(1)
	
	browser.find_by_text(u"修改").first.click() 
	time.sleep(1)
	
	
	
	#跳转到对应iframe
	with browser.get_iframe('layui-layer-iframe1') as iframe1:
		
		
		#因为页面有滚动才能显示，设置div的属性
		browser.evaluate_script('document.getElementsByClassName("warp_scroll_body warp_scroll_box")[0].style="overflow:scroll;"')
		#利用js让页面滚动，显示被遮挡部分
		browser.execute_script('$(".warp_scroll_body warp_scroll_box").scrollTop(500)')
	
		#
		#time.sleep(3)
		browser.fill("dflzOrganization.remark2",code)
		


	
	browser.find_by_text(u"保存并关闭").click()  
	browser.find_by_text(u"确定").click() 		   

		
		
#读取excel 	
def readExcl():
	 xl=xlrd.open_workbook(excPath)
	 table=xl.sheets()[0]
	 
	 #获取excel的列，存入list
	 zhi=table.col_values(0)
	 keys=table.col_values(1)
	 print(len(keys))
	 print(len(zhi))
	
	#遍历list，并自动录入，并保存字典
	 for i in range(1,len(keys)):
			dlinput(keys[i],zhi[i])
			dic[keys[i]]=zhi[i]	
			
			
	 """
	 clos=table.col_values(0)
	 clos2=table.col_values(1)
	 print(clos)
	 print(clos2)

	 for i in range(1,len(clos)):
	
		  dic[clos2[i]]=clos[i]
	 """

	

def main():
	login()
	browser.visit(dlurl)#界面跳转
	dlinput()
	readExcl()

	
if __name__ == '__main__':
	 main()
			
	