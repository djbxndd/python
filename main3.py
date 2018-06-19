## 品质报表自动化处理程序(版本3.0）
# （版本1.0 main.py 20170825）；利用双控变量控制单元格的读写（核心代码）
#  （版本2.0 main2.py 20170925);利用pandas库对数据进行处理；只处理一天的数据。
#  (版本3.0 main3.py 20171002) 处理的数据源是从某月连续多天的数据。即“大表”思维；
## 修改履历：在第一个模块SingleFile（）中对表头进行规范化处理；
#-------------------------------------------------------------------------------
## 加载必要的库；
##import os
import datetime as dt
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

import xlrd,sys
import xlwt
from datetime import date,datetime

#-------------------------------------------------------------------------------
'''
模 块1：文件数据预处理（剔除多余的行与列）；
函数名：SingleFile(path,filename,result)
参 数1：path=文件路径
参 数2：filename=文件名；
参 数3：result=输出结果指定的文件名；

函数返回值：处理结果的文件名（由主程序指定）

'''
#-------------------------------------------------------------------------------
def SingleFile(path,filename,result):
    p1=path
    f1=filename
    Output_name=result
    print("当前正在处理的文件是{}目录下的{}文件，请核对！".format(p1,f1))
    
    workbook=xlrd.open_workbook(f1)   #利用xlrd库打开EXCEL文件；
    print(workbook.sheet_names())  #输出这个工作薄文件所有的工作表名；
    #获取第一张工作表的表名\行数\列数
    sheet= workbook.sheet_names()[0]  #注意，如果有多张sheet,引用时要变换下标；
    sheet1 = workbook.sheet_by_name(sheet)
    nName=sheet1.name     #工作表名字：全局变量
    nRows=sheet1.nrows    #总行数（注意：没有计算空白行）全局变量
    nCols=sheet1.ncols    #总列数；全局变量
    print(nName,nRows,nCols)  #调试用；

    #准备写入的工作薄和工作表文件；
    workbook=xlwt.Workbook(encoding='utf-8')  #创建工作簿
    sheet2=workbook.add_sheet("sheet1",cell_overwrite_ok=True)  #创建sheet
##    Output_name=p1+'result.xls'  #是否只能输出Excel2003格式？【如果文件存在的话，程序直接覆盖原文件】
    
    #初始化单元格内容的字体样式
    style=xlwt.XFStyle()
    font=xlwt.Font()    #设置单元格内字体的式样；
    font.name="Times New Roman"
    font.bold=False
    #设置样式的字体
    style.font = font
    
    ##I,J 源数据文件指针；
    ##P,Q 目标文件指针
    ##经验分享，（1）设置两个文件指针，（2）动态打印变量的值，方便查找出错原因。搞定。
    ##完成日期：2017/8/26

    # 定义符合条件的列的索引列表；(以下为熊小苗所需字段的索引值）
    ##field_index=[0,1,3,4,6,22,29,32,33,34,35,36,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87]
    # 以下为熊琼B社数据源所需字段；
    field_index=[0,1,3,4,13,19,20,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76]
    
    i=5;p=0   ##从表头行开始；   
    while i<nRows:  #外循环，从第5行开始到最大行数；（注意边界是否包含？）
        j=0;q=0
        for k in range(0,len(field_index)):  #内循环开始，处理一行的指定列；
            j=field_index[k]   
            CellValue=sheet1.row(i)[j].value    #读数据给变量；        
            if j==0 and (CellValue.startswith("SUB") or CellValue.startswith("TOTAL")) : #如果单元格的值是"SUB"或者"TOTAL"，代表此行是汇总行，剔除掉；
                p=p-1
##                print("源数据的汇总行是：{}  目标数据行的当前行是：{}".format(i,p),end=" ")
##                print()
                break    
            else: 
                sheet2.write(p,q,CellValue,style)   #将变量值写入新表指定的单元格,同时设定字体；         
            j+=1;q+=1
        i+=1;p+=1
##        print();
    workbook.save(Output_name)  #保存文件(为什么如果输出文件名不是全局变量，数据无法保存）；
##------------------------------
    ## 增加处理表头的功能；重新定义，值由主模块全局变量传入；修改日期20170928  变更后生成的EXCEL表格多了一个索引列；对下一个模块的程序是否有影响；
    rbb_pd=pd.read_excel(Output_name,sheetname="sheet1",header=0,skiprows=None)
    rbb_pd.columns=heading
    rbb_pd.to_excel(Output_name,"sheet1")   ##将重新定义了表头的文件存盘；
    return Output_name
#-------------------------------------------------------------------------------
'''
模块2：数据深加工处理
函数名：FileProcess(path,filename,flhzb,fenlei)
参数1：path=文件路径
参数2：filename=文件名；
参数3：flhzb=输出结果指定的文件名；
参数4：fenlei=品目分类表文件名（要确保此数据为最新状态，如果不是最新的，必须按照数据目录下的格式进行更新）

函数返回值：处理结果的文件名（由主程序指定）
版本2的 局限性：当前仅限于处理某一天的日报，如果是多天的日报，那么对于【REPORT DATE】列的取数要重新设计；
版本3：算法已修改，处理连续多天的数据。直接利用数据透视表方法（分类为多个字段）
'''
#-------------------------------------------------------------------------------
## 模块2_excel数据深度加工由pandas 库实现；
## 修改说明：程序功能变更为处理连续多日的日报表数据，直接使用Pandas库实现。
def FileProcess(path,filename,flhzb,fenlei,hzb):
    p1=path
    rbb=filename    
    hz=hzb
    print("当前正在对{}目录下的{}文件进行分类汇总......".format(p1,rbb))  
    
    src_file=rbb
    FL=fenlei
    hzzb=flhzb
    
    rbb_pd=pd.read_excel(src_file,sheetname="sheet1")
  
    FL_pd=pd.read_excel(FL,sheetname="FPC_SMT_LINK")   #此处指定工作表的名字的方法不好，日后代码维护困难！
    
    out_pd=pd.merge(rbb_pd,FL_pd,on='ITEM 6',how='left')   # 完成日报表【分类字段】和【...】的关联更新

    out_pd=out_pd.sort_values(["riqi","fenlei","FPC_CPN","ITEM 6"],ascending=[True,False,True,True])   ##对统计结果按四列进行排序{日期、分类【一贯非一贯】、系列（A61..)，再按LOT，降|升|升序}

    Fpc_NG_sum=[0]*len(out_pd)   # 定义FPC前33项不良合计的列表；先置空值，再用循环的方法对逐行求和；

    for i in range(1,len(out_pd)):    # 模块1传入的源文件多了一个索引列，对此段代码有无影响，需要再验证；20171002
        Fpc_NG_sum[i-1]=out_pd.iloc[i-1,7:40].sum()
        
    out_pd['SumFpcNG']=Fpc_NG_sum               # 求每一个LOT FPC前工程33项不良之和
    
    out_pd['QtyOut']=out_pd["QTY IN"]-out_pd['SumFpcNG']   # 每一个LOT的产出；
    
    out_pd['lotOK']=out_pd['QtyOut']/out_pd["QTY IN"]   # 每一个LOT的良率；

    out_pd.to_excel(hz,"sheet1")   # 将结果输出到EXCEL文件

##    hzb_f3=out_pd.pivot_table(index=["fenlei","FPC_CPN"],columns="riqi",values=["QTY IN","QtyOut","lotOK"],aggfunc=sum);  # aggfunc=sum表示求和；
## 
##    hzb_f3=hzb_f3.stack()   #stack()|unstack() 可以改变数据透视表排列方式；UnStack()不好看；

    hzb_f3=out_pd.pivot_table(index=["fenlei","FPC_CPN",],columns="riqi",values=["QTY IN","QtyOut","lotOK"],aggfunc=sum);  # aggfunc=sum表示求和；
  
##    hzb_f3=out_pd.pivot_table(index=["fenlei","FPC_CPN"],columns="riqi",values=["QTY IN"],aggfunc=sum);  # aggfunc=sum表示求和；

    hzb_f3.to_excel(hzzb,"sheet1")   #将分类的结果输出到EXCEL文件

    return FLhzb,hzb
    

##  下面的代码已经作废。（改用上面的方法，利用数据透视表函数一步到位；

##    hzb_fl=out_pd.groupby('FPC_CPN').size()   #指定一个字段【fpc系列统计】（即A61|A71....)进行分类汇总（统计符合条件的行数）；
##
##    hzb_f3=out_pd.groupby('FPC_CPN').sum()   #指定所有数字型字段求和；
##
##    hzb_f3['XlOK']=hzb_f3['QtyOut']/hzb_f3["QTY IN"]    #重新求一次系列的良率
##
##    xlMax=out_pd.groupby('FPC_CPN')['lotOK'].max()        #求某一天的系列的良率最大值与最小值
##    xlMin=out_pd.groupby('FPC_CPN')['lotOK'].min()
##
##    hzb_f3['riqi']=riqi     #为DataFrame对象的某一列赋新值【riqi变量为全局变量，由主程序传入】
##
##    hzb_f3['SEQ']='#N/A'
##    hzb_f3['Lot']='#N/A'   
##    return hzb

#-------------------------------------------------------------------------------

##主程序开始；


## 文件检测模块（暂未编写）；
##print("文件自动检测开始.....")

fstatus=True

if fstatus:    ##此处以后替换成异常捕获处理程序；
    print("本系统（测试版）功能说明：只能处理【单一文件数据预处理】--【数据统计与结果输出】")    
    print("                     欢迎使用本系统！                  ")
    print("=======================================================")
    print()
    
    path=input("请输入日报文件的路径(注意:目录中以'\\'结尾):")
    fname=input("请输入要处理的文件名【强制规定：文件名的命名规则2017mmdd.xlsx(mm=两位的年份；dd=两位的月份）】：")

    ##os.chdir(path)   ## 设置工作目录
    riqi=fname.split(sep=".")[0]  ## riqi全局变量的作用是：最终处理结果文件的日期列存储值，取文件主名；

    out=path +"result.xls"      ## 预处理结果文件；
    hzb=path +"hzb.xlsx"        ## 输出【日报表与分类表】关联后的结果；
    FLhzb=path+"FLhzb.xlsx"          ## 输出按【系列】分类汇总后的结果；

    fenlei=path+"fenlei.xlsx"        #分类表；

    headstru=path+"datahead.xlsx"    ##日报表表头定义文件【工作目录下的datahead.xlsx】；

    heading=pd.read_excel(headstru,sheetname=0,header=0,skiprows=None)["newname"]   ## 定义表头的全局变量；
    
    print("数据预处理开始，请稍等......................................")
    
    output=SingleFile(path,fname,out)  ##调用单一文件处理函数，对数据进行预处理；
    print("数据预处理完成！请到{}下核对处理结果是否正确".format(output))
    print()

    print("数据深度加工开始.....本次将生成按系列的不良率报表数据！")
    xlOK=FileProcess(path,out,FLhzb,fenlei,hzb)
    
    print("数据预处理完成！请到{}下核对处理结果是否正确".format(xlOK))
    
    print()
    print("===================== 谢 谢 使 用 本 系 统！ =======================")
    
    
else:
    print("您输入的文件名有误，或者文件不存在，请先确认！")

#-------------------



