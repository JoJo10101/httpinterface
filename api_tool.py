#coding:utf-8
import xlrd
import urllib2,urllib
import traceback,sys
from xlutils import copy


dic={} #存放执行结果

dic1={} #存放判断结果

wkb=xlrd.open_workbook('data.xls') #读取接口文档

sheet=wkb.sheets()[0] 

basis='http://'+wkb.sheet_names()[0]
#basis='http://10.37.18.91:8250'

nrows=sheet.nrows  #获取行数


for i in range(1,nrows): #循环取excel行内容
    
    api=sheet.cell(i,1).value #获取excel‘接口’内容

    url=basis+api 

    #value=sheet.cell(i,5).value.encode('utf-8') #获取excel‘参数’内容

    #data=eval('dict(%s)'%value)

    #fix meilin
    value=sheet.cell(i,5).value.encode('utf-8')
    data=eval(value)

    data=urllib.urlencode(data)  

    #value1=sheet.cell(i,2).value.encode('utf-8') #获取excel‘请求头’内容
    
    #head=eval('dict(%s)'%value1)
    #fix meilin
    value1=sheet.cell(i,5).value.encode('utf-8')
    head=eval(value1)

    value2=sheet.cell(i,7).value #获取excel'期望数据'内容

    if value1 == '': #判断excel中是否存在“请求头”

        url2=urllib2.Request(url,data) #发送请求 

    else:

        url2=urllib2.Request(url,data,head)  #加请求头发送请求
    
    try :

        response=urllib2.urlopen(url2)  #获取响应内容

        apicontent=response.read()  

        content=apicontent.decode('utf-8')

        dic[i]=content

        if value2 in content:  #响应结果与excel中“预期结果”做比对
            
            result='pass'
            
        else:
            
            result='failed'

        dic1[i]=result

    except:  
        
        exc="".join(traceback.format_exception(*sys.exc_info()))

        a=exc.decode('utf-8','ignore')

        dic[i]=a       

        dic1[i]='failed'

wkb_cp=copy.copy(wkb) 

sheet1=wkb_cp.get_sheet(0)

for j in dic: #将执行结果存放excel中
    
    sheet1.write(j,6,dic[j])

col=0

for k in dic1: #将判断结果存放excel中

    sheet1.write(k,8,dic1[k])

    if dic1[k]=='pass':
        
        col=col+1

wkb_cp.save('result.xls')

icol=col*100/i

print u'运行结束'

print u'共执行：%d次'%i

print u'通过率：%d'%icol+'%'

raw_input()
