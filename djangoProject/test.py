import datetime
import json
import os
import re
import time
import urllib.parse

import requests
import pandas as pd
from docx import Document
from openpyxl.styles import Border, Side, Alignment, Font

from sqlalchemy import create_engine

from djangoProject.color import mos_wirte_excel

engine = create_engine("mysql+pymysql://root:123456@127.0.0.1:3306/wubian")

from djangoProject.erpApi import requesterp
from djangoProject.lingxingApi.aes import md5_encrypt, aes_encrypt
import orjson
from djangoProject.lingxingApi.http_util import HttpBase
from djangoProject.lingxingApi.resp_schema import ResponseResult
from djangoProject.lingxingApi.sign import SignBase
# a = requesterp()
def listing_price_load():
    df = pd.read_excel(r"C:\Users\wb\Desktop\Listing20230303-490113876584333312.xlsx").fillna("")
    df = df.rename(columns={"店铺":"store","ASIN":"asin","父ASIN":"parent_asin","MSKU":"msku","SKU":"GS-sku","价格":"price"},inplace=False)
    df = df[['store','asin','parent_asin','msku','GS-sku','price']]
    df["date"] = datetime.datetime.now()
    df["sku"] = df["GS-sku"].replace("GS-","")
    df["update_date"] = datetime.datetime.now()
    df["create_date"] = datetime.datetime.now()
    df.to_sql('listing_price_date',engine,chunksize=10000,if_exists='append',index=False)
    print("done")


# listing_price_load()


def api_listing_price_load():
    df_sid = pd.DataFrame()
    sids = requests.get("http://43.142.117.35/get_store_id")
    data = eval(sids.text)
    for i in data:
        sid = i.get("sid")
        name = i.get("name")
        dftmp = pd.DataFrame({"sid":sid,"name":name},index=[0])
        df_sid = df_sid.append(dftmp)
    df_sid = df_sid.reset_index()
    del df_sid["index"]
    a=11
    a = requests.get("http://43.142.117.35/get_listing",params={"offset":0})
    data = eval(a.text)
    data = data.get("data")
    df = pd.DataFrame()
    for i in data:
        status = i.get("status")
        is_delete = i.get("is_delete")
        if is_delete == 0 and status == 1:
            asin = i.get("asin")
            parent_asin = i.get("parent_asin")
            msku = i.get("seller_sku")
            sid = i.get("sid")
            GS_sku = i.get("local_sku")
            sku = GS_sku.replace("GS-","")
            price = i.get("price")
            dftmp = pd.DataFrame({"sid":sid,"asin":asin,"parent_asin":parent_asin,"msku":msku,"status":status,"GS-sku":GS_sku,"sku":sku,"price":price,"is_delete":is_delete},index=[0])
            df = df.append(dftmp)
    df = df.reset_index()
    del df['index']
    df = pd.merge(df,df_sid,left_on=['sid'],right_on=['sid'],how="left").reset_index()
    del df['index']
    df["date"] = datetime.datetime.now()
    df["update_date"] = datetime.datetime.now()
    df["create_date"] = datetime.datetime.now()
    df = df.rename(
        columns={"name": "store"},
        inplace=False)
    df = df[['store', 'asin', 'parent_asin', 'msku', 'GS-sku','sku','price','date','update_date','create_date']]
    df.to_sql('listing_price_date', engine, chunksize=10000, if_exists='append', index=False)
    # df.to_excel(r"C:\Users\wb\Desktop\listing.xlsx")


# api_listing_price_load()
# start_date = '2023-01-01'
# end_date = '2023-01-31'
# req = {
#     "start_date":start_date,
#     "end_date":end_date,
# }
# a = requests.post("http://43.142.117.35/old_profitSettlement",params=req)
# a=11
req_body = {
    "sid": 2613,
    "offset": 0,
    "length": 1000,
    "sort_type": "desc",
    "sort_field": "volume",
    "start_date": "2023-02-13",
    "end_date": "2023-02-13",
    "summary_field": "msku",
    }

# a = requests.post("http://43.142.117.35/get_fbawarehouse")
# a = requests.post("http://43.142.117.35/productPerFormance",params=req_body)
# b = eval(json.dumps(a.text))
# eval(b)
# a.json()
a=1

#
def 箱标():
    import winreg
    def desktop_path():
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        desktop = winreg.QueryValueEx(key, "Desktop")[0]
        print(desktop)
        return desktop
    from docx import Document
    def delete_paragraph(paragraph):
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None
    desktop_path = desktop_path()
    doc = Document(desktop_path+"\日本海外仓箱标\日本海外仓箱标模板.docx")
    df = pd.read_excel(desktop_path+"\日本海外仓箱标\需生成的数据文件 - 副本.xlsx")
    df_data = pd.DataFrame()
    for i in df.index:
        xiangshu = int(df.loc[i,"箱数"])
        sku = df.loc[i,"品名"]
        num = int(df.loc[i,"单箱数量"])
        zimu = df.loc[i,"字母标签"]
        code = df.loc[i,"箱子编号"]
        if xiangshu>1:
            df_tmp = pd.DataFrame({"品名":sku,"箱子编号":code,'字母标签':zimu,"单箱数量":num},index=[0])
            for n in range(0,xiangshu):
                df_data = df_data.append(df_tmp)
        else:
            df_tmp = pd.DataFrame({"品名": sku, "箱子编号": code, '字母标签': zimu, "单箱数量": num}, index=[0])
            df_data = df_data.append(df_tmp)
    df_data = df_data.reset_index()
    sku_list = df_data['品名'].tolist()
    num_list = df_data['单箱数量'].tolist()
    zimu_list = df_data['字母标签'].tolist()
    code_list = df_data['箱子编号'].tolist()
    list1 = doc.paragraphs
    for i in list1:
        delete_paragraph(i)
        if (len(doc.paragraphs))==len(sku_list):
            break
    # doc.save(desktop_path+"\日本海外仓箱标\日本海外仓箱标.docx")
    doc.save(r'C:\Users\wb\Desktop\test.docx')

    doc = Document(r'C:\Users\wb\Desktop\test.docx')
    children = doc.element.body.iter()
    count = 0  # 写一个count是为了，可以定位是哪个文本框，因为我用索引失败了
    for child in children:
        # 通过类型判断目录
        if child.tag.endswith('txbx'):
            count += 1
            if count == 3:
                for ci in child.iter():
                    if ci.tag.endswith('main}r'):
                        if ci.text == '型号：':
                            ci.text = ''
                        if ci.text == '字母区分：':
                            ci.text = ''
                        if ci.text == '数量：':
                            ci.text = ''
                        if ci.text == '箱子编号：':
                            ci.text = ''
    doc.save(r'C:\Users\wb\Desktop\test.docx')
    doc = Document(r'C:\Users\wb\Desktop\test.docx')
    sku_dict = {}
    for i in range(1,(len(sku_list)*2)+1):
        if i==1:
            sku_dict[1] = 1
        elif i%2==0:
            sku_dict[i] = sku_dict[i-1]
        else:
            sku_dict[i] = sku_dict[i-1] + 1
    m = 1
    n = 0
    children = doc.element.body.iter()
    for child in children:
        # 通过类型判断目录
        if child.tag.endswith('txbx'):
            for ci in child.iter():
                if ci.tag.endswith('main}r'):
                    if ci.text == '型号：':
                        # print(ci.text)
                        ci.text = "型号：%s"%sku_list[n]
                    if ci.text == '字母区分：':
                        # print(ci.text)
                        ci.text = "字母区分：%s" % zimu_list[n]
                    if ci.text == '数量：':
                        # print(ci.text)
                        ci.text = "数量：%s" % num_list[n]
                    if ci.text == '箱子编号：':
                        # print(ci.text)
                        ci.text = "箱子编号：\n%s"%code_list[n]
                        n+=1
                print("正在生成")
        doc.save(r'C:\Users\wb\Desktop\test.docx')

#

# 箱标()

# for index,sku in enumerate(sku_list):




def 补货表生成V1():

    #sales and traffic 数据透视处理
    df_7D = pd.read_excel(r'C:\Users\wb\Desktop\工作簿2.xlsx',sheet_name="7天")
    df_7D = df_7D.pivot_table(index=['（子）ASIN'],values=['会话次数 – 总计','会话次数 – 总计 – B2B','订单商品总数','订单商品总数 - B2B'],aggfunc=sum).reset_index()
    df_14D = pd.read_excel(r'C:\Users\wb\Desktop\工作簿2.xlsx',sheet_name="14天")
    df_14D = df_14D.pivot_table(index=['（子）ASIN'],values=['会话次数 – 总计','会话次数 – 总计 – B2B','订单商品总数','订单商品总数 - B2B'],aggfunc=sum).reset_index()
    df_magInv = pd.read_excel(r'C:\Users\wb\Desktop\工作簿2.xlsx',sheet_name="管理亚马逊库存")
    df_asin_to_sku = df_magInv[['sku','asin']]
    df_7D = pd.merge(df_7D,df_asin_to_sku,left_on=['（子）ASIN'],right_on=['asin'],how='left')
    df_14D = pd.merge(df_14D,df_asin_to_sku,left_on=['（子）ASIN'],right_on=['asin'],how='left')
    del df_7D["asin"]
    del df_14D["asin"]
    df_7D['总流量'] = df_7D['会话次数 – 总计'] + df_7D['会话次数 – 总计 – B2B']
    df_7D['总出单'] = df_7D['订单商品总数'] + df_7D['订单商品总数 - B2B']
    df_7D['转化率'] = df_7D['总出单'] / df_7D['总流量']
    df_14D['总流量'] = df_14D['会话次数 – 总计'] + df_14D['会话次数 – 总计 – B2B']
    df_14D['总出单'] = df_14D['订单商品总数'] + df_14D['订单商品总数 - B2B']
    df_14D['转化率'] = df_14D['总出单'] / df_14D['总流量']
    df_7D['转化率'] = df_7D['转化率'].map(lambda x: format(x,'.2%'))
    df_14D['转化率'] = df_14D['转化率'].map(lambda x: format(x,'.2%'))



    #补货表数据处理
    df_buhuo = pd.read_excel(r'C:\Users\wb\Desktop\工作簿2.xlsx',sheet_name="补货")
    df_buhuo['预留'] = df_buhuo['FC transfer'] + df_buhuo['FC Processing']

    #管理亚马逊库存
    df_magInv = pd.merge(df_magInv,df_buhuo[['ASIN','预留']],left_on=['asin'],right_on=['ASIN'],how="left")
    df_magInv['IN-STOCK库存'] = df_magInv['afn-fulfillable-quantity']
    df_magInv['总库存'] = df_magInv['预留'] + df_magInv['afn-fulfillable-quantity'] +df_magInv['afn-inbound-working-quantity'] + df_magInv['afn-inbound-shipped-quantity'] + df_magInv['afn-inbound-receiving-quantity']


    #库存明细生成
    df_Invdetail = df_asin_to_sku
    df_Invdetail.rename(columns={"sku": "MSKU"},inplace=True)
    df_basedata = pd.read_excel(r'C:\Users\wb\Desktop\补货表基础信息.xlsx',sheet_name="基础信息")
    df_Invdetail = pd.merge(df_Invdetail,df_basedata,
                            left_on=['MSKU'],right_on=['MSKU'],how="outer").fillna(0)


    df_Invdetail = pd.merge(df_Invdetail,df_magInv[['sku','IN-STOCK库存','总库存','your-price']],left_on=['MSKU'],right_on=['sku'],how="left").fillna("")
    df_Invdetail.rename(columns={"总库存": "亚马逊总库存"},inplace=True)
    df_Invdetail['FBA在库预估'] = df_Invdetail['亚马逊总库存']
    df_Invdetail['海外仓库存'] = 0
    df_Invdetail['预创货件'] = 0
    df_Invdetail = pd.merge(df_Invdetail,df_7D[['sku','总流量','转化率','总出单']],left_on=['sku'],right_on=['sku'],how="left")
    df_Invdetail.rename(columns={"总流量": "7天流量","转化率":"7天转化率","总出单":"7天订单数量"},inplace=True)
    df_Invdetail['7天实际日均'] = round(df_Invdetail['7天订单数量']/7,1)
    df_Invdetail['常规补货计划日均'] = 0
    df_Invdetail = pd.merge(df_Invdetail,df_14D[['sku','总流量','转化率','总出单']],left_on=['sku'],right_on=['sku'],how="left")
    df_Invdetail.rename(columns={"总流量": "14天流量","转化率":"14天转化率","总出单":"14天订单数量"},inplace=True)
    df_Invdetail['14天日均数量'] = round(df_Invdetail['14天订单数量']/14)
    df_Invdetail['7/14增长'] = (df_Invdetail['7天实际日均']-df_Invdetail['14天日均数量'])/df_Invdetail['14天日均数量']
    df_Invdetail.rename(columns={"asin": "ASIN","your-price":"价格"},inplace=True)
    df_Invdetail = df_Invdetail.fillna(0)
    weight_av = 6000
    rate = 6.6
    canshu = 7
    df_Invdetail['计费重'] = round((df_Invdetail['长']*df_Invdetail['宽']*df_Invdetail['高'])/weight_av)
    df_Invdetail['价格'] = df_Invdetail['价格'].replace("",0)
    df_Invdetail["$头程"] = round(df_Invdetail['计费重']*canshu/rate,1)

    df_Invdetail['分割1'] = ""
    df_Invdetail['分割2'] = ""
    df_Invdetail['3'] = ""
    df_Invdetail['4'] = ""
    df_Invdetail['5'] = ""
    df_Invdetail['6'] = ""


    df_Invdetail = df_Invdetail[["MSKU","公司SKU","数据","父体","X备注","备注","分割1","颜色大类",
                                 "细分颜色","组合方式","材质","可扩展",
                                 "20寸的尺寸","ASIN","分割2","IN-STOCK库存","FBA在库预估",
                                 "亚马逊总库存","3","海外仓库存","预创货件","4","7天流量","7天转化率",
                                 "7天订单数量","7天实际日均","常规补货计划日均","5","14天订单数量",
                                 "14天日均数量","7/14增长","14天流量","14天转化率","6","佣金","FBA费",
                                 "成本","高","长","宽","计费重","$头程","价格"]]
    df_Invdetail['定价毛利率'] = ((df_Invdetail['价格']-df_Invdetail['价格']*df_Invdetail['佣金']-df_Invdetail['FBA费']-(df_Invdetail['成本'])/rate)-df_Invdetail['$头程'])/df_Invdetail['价格']
    df_Invdetail['DEAL'] = 0
    df_Invdetail['Coupon'] = 0
    df_Invdetail['Prime'] = 0
    df_Invdetail['OFF'] = 0
    df_Invdetail['售价'] = df_Invdetail['价格'] * (1-df_Invdetail['OFF'])-df_Invdetail['DEAL']*df_Invdetail['价格']-df_Invdetail['Coupon']-df_Invdetail['Prime']*df_Invdetail['价格']
    df_Invdetail['毛利率'] = (df_Invdetail['售价']-df_Invdetail['售价']*df_Invdetail['佣金']-df_Invdetail['成本']/rate-df_Invdetail['FBA费']-df_Invdetail['$头程'])/df_Invdetail['售价']
    df_Invdetail['毛利额$'] = round(df_Invdetail['售价'] * df_Invdetail['毛利率'],2)
    df_Invdetail['7'] = ""
    df_Invdetail['INSTOCK库存售罄天数'] = "=IFERROR(P2/AA2,0)"
    df_Invdetail['FBA预估库存售罄天数'] = "=IFERROR(Q2/AA2,0)"
    df_Invdetail['FBA总库存（含在途）售罄天数'] = "=IFERROR(R2/AA2,0)"
    df_Invdetail['INSTOCK差额＜60'] = "=BC2-BB2"
    df_Invdetail['实际售罄天数'] = "=IFERROR(R2/Z2,0)"
    df_Invdetail['8'] = ""
    df_Invdetail['9'] = ""
    df_Invdetail['安全天数'] = ""
    df_Invdetail['应补货'] = "=IF(AA2*BH2-R2>0,AA2*BH2-R2,0)"
    df_Invdetail['备注2'] = ""
    df_Invdetail['1月待安排'] = ""
    df_Invdetail['10'] = ""
    df_Invdetail["15天特批"] = ""
    df_Invdetail["25天"] = ""
    df_Invdetail["35天"] = ""
    df_Invdetail["45天"] = ""
    df_Invdetail["60天"] = ""
    df_Invdetail["合计"] = "=SUM(BM2:BQ2)"
    df_Invdetail["复核天数"] = "=IFERROR(BR2/AA2,0)"
    df_Invdetail["合计天数"] = "=IFERROR(BC2+BS2,0)"
    df_Invdetail["差值"] = "=IFERROR(BR2-BI2,0)"
    df_Invdetail['11'] = ""
    df_Invdetail["工厂待出货数量"] = ""
    df_Invdetail["预下单量"] = ""
    df_Invdetail['12'] = ""
    df_Invdetail["x备注"] = ""
    df_Invdetail["F"] = ""
    df_Invdetail['定价毛利率'] = df_Invdetail['定价毛利率'].map(lambda x: format(x,'.2%'))
    df_Invdetail['佣金'] = df_Invdetail['佣金'].map(lambda x: format(x,'.2%'))
    writer = pd.ExcelWriter(r"C:\Users\wb\Desktop\LXGG.xlsx", engine='openpyxl')  # 创建数据存放路径
    df_Invdetail.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()  # 文件保存
    writer.close()

    a=11



# 补货表生成V1()

# a=11


def 供应商问题标签():
    df = pd.read_excel(r'C:\Users\wb\Desktop\拉杆箱次品原因汇总.xlsx',sheet_name="Sheet2")

    for i in df.index:
        list1 = df.loc[i,'汇总'].split("，")
        list_name = df.columns.values.tolist()
        for j in list1:
            if j.strip() not in list_name:
                df.loc[i,'%s'%j.strip()] = 1
            else:
                df.loc[i, '%s' % j.strip()] = 1
        a=1

    df=df.fillna(0)
    df.to_excel(r"C:\Users\wb\Desktop\123.xlsx")
    a=11



    import re
    df = pd.read_excel(r'C:\Users\wb\Desktop\拉杆箱次品原因汇总.xlsx',sheet_name="Sheet1")
    for i in df.index:
        list1 = df.loc[i,'汇总'].split("，")
        list_name = df.columns.values.tolist()
        for j in list1:
            if j.strip() not in list_name:
                if j == "内装折皱批量":
                    df.loc[i, '内装折皱'] = "批量"
                elif j == "三角不平批量":
                    df.loc[i, '三角不平'] = "批量"
                else:
                    try:
                        num = re.search("\d+",j)
                        num = int(num.group(0))
                        name = j.replace("%s个"%num,"")
                        name = name.replace("个%s" %num,"").strip()
                        df.loc[i,'%s'%name] = num
                    except:
                        df.loc[i, '%s'%j.strip()] = 1
            else:
                try:
                    num = re.search("\d+", j)
                    num = int(num.group(0))
                    name = j.replace("%s个" % num, "")
                    name = name.replace("个%s" % num, "")
                    df.loc[i, '%s' % name] = num
                except:
                    pass

        a=1

    df.to_excel(r"C:\Users\wb\Desktop\234.xlsx")



    df_1 = pd.read_excel(r'C:\Users\wb\Desktop\123.xlsx')
    df_2 = pd.read_excel(r'C:\Users\wb\Desktop\234.xlsx')

    list1 = df_1.columns.values.tolist()
    list2 = df_2.columns.values.tolist()



    df = df_1.append(df_2)
    df.to_excel(r"C:\Users\wb\Desktop\aaaa.xlsx")

    #
    # a=11

#
#
def write_excel():
    Inv_arr = ["MSKU","公司SKU","数据","父体","X备注","备注","分割1","颜色大类",
                                     "细分颜色","组合方式","材质","可扩展",
                                     "20寸的尺寸","ASIN","分割2","IN-STOCK库存","FBA在库预估",
                                     "亚马逊总库存","3","海外仓库存","预创货件","4","7天流量","7天转化率",
                                     "7天订单数量","7天实际日均","常规补货计划日均","5","14天订单数量",
                                     "14天日均数量","7/14增长","14天流量","14天转化率","6","佣金","FBA费",
                                     "成本","高","长","宽","计费重","$头程","价格","定价毛利率","DEAL",
                "Coupon","Prime","OFF","售价","毛利率","毛利额$","7","定价毛利率",
               "INSTOCK库存售罄天数","FBA预估库存售罄天数",
               "FBA总库存（含在途）售罄天数","INSTOCK差额＜60",
               "实际售罄天数","8","9","安全天数","应补货","备注2","1月待安排","10","备注"
               "15天特批","25天","35天","45天","60天","复核天数",
               "合计天数","差值","11","工厂待出货数量","预下单量","12","x备注","F"]
    inv_dic = {}
    title = Inv_arr
    import openpyxl
    excel=openpyxl.load_workbook(r'C:\Users\wb\Desktop\LXGG.xlsx')
    sheet=excel['Sheet1']
    sheet.freeze_panes='AB2' #冻结单元格
    zm = [chr(i) for i in range(ord("A"), ord("Z") + 1)]
    for t in title:
        if title.index(t) == 26:
            zm.extend(["A%s" % i for i in zm])
        if title.index(t) == 53:
            zm.extend(["B%s" % i for i in zm])
        sheet[f"{zm[title.index(t)]}1"].value = t
        inv_dic.update({t: zm[title.index(t)]})  # {"标题": 列序号}


    border = Border(left=Side(border_style='thin', color='000000'),
                    right=Side(border_style='thin', color='000000'),
                    top=Side(border_style='thin', color='000000'),
                    bottom=Side(border_style='thin', color='000000'))

    #列染色和设置分割列宽,type=1为分割列
    def col_colors(names,color,type=0):
        from openpyxl.styles import PatternFill
        for name in names:
            col = sheet[inv_dic['%s'%name]]
            fills = PatternFill("solid", fgColor=color)
            if type == 1:
                sheet.column_dimensions[inv_dic["%s"%name]].width = 0.46
            for cell in col:
                cell.fill = fills

    sheet.column_dimensions.group(inv_dic['公司SKU'],inv_dic['父体'], hidden=True)
    sheet.column_dimensions.group(inv_dic['颜色大类'],inv_dic['ASIN'], hidden=True)
    sheet.column_dimensions.group(inv_dic['佣金'],inv_dic['$头程'], hidden=True)
    sheet.column_dimensions.group(inv_dic['DEAL'],inv_dic['毛利额$'], hidden=True)
    sheet.column_dimensions.group(inv_dic['备注2'],inv_dic['1月待安排'], hidden=True)
    # sheet.column_dimensions.outline_level=None

    col_colors(['分割1','分割2','3','4','5','6','7','8','9','10','11','12'],"C00000",type=1)
    col_colors(['FBA在库预估'],"B4C6E7")
    col_colors(['海外仓库存','预创货件'],"FFF2CC")
    col_colors(['7天流量','7天转化率'],"D6DCE4")
    col_colors(['常规补货计划日均'],"F8CBAD")
    col_colors(['INSTOCK库存售罄天数','FBA预估库存售罄天数','FBA总库存（含在途）售罄天数','INSTOCK差额＜60'],"D6DCE4")
    col_colors(["14天流量","14天转化率","佣金","FBA费","成本","高","长","宽","计费重","$头程"],"D6DCE4")
    col_colors(["价格","定价毛利率"],"E2EFDA")
    col_colors(["实际售罄天数"],"FCE4D6")
    col_colors(["安全天数"],"DDEBF7")
    col_colors(["应补货"],"F4B084")
    col_colors(["备注2","1月待安排"],"FFF2CC")
    col_colors(["25天","35天"],"DDEBF7")
    col_colors(["45天","60天"],"FCE4D6")
    col_colors(["复核天数"],"FFE699")
    col_colors(["合计天数"],"DDEBF7")
    font1 = Font(
        name="等线",   # 字体
        size=11,         # 字体大小
        color="000000",  # 字体颜色，用16进制rgb表示
        bold=True,       # 是否加粗，True/False
        italic=False,     # 是否斜体，True/False
        strike=None,     # 是否使用删除线，True/False
        underline=None,  # 下划线, 可选'singleAccounting', 'double', 'single', 'doubleAccounting'
    )
    font2 = Font(
        name="等线",   # 字体
        size=11,         # 字体大小
        color="000000",  # 字体颜色，用16进制rgb表示
        bold=False,       # 是否加粗，True/False
        italic=False,     # 是否斜体，True/False
        strike=None,     # 是否使用删除线，True/False
        underline=None,  # 下划线, 可选'singleAccounting', 'double', 'single', 'doubleAccounting'
    )



    align=Alignment(horizontal='center',vertical='center')
    for i in range(1,sheet.max_row+1): #遍历行号
        for j in range(1,sheet.max_column+1): # 遍历当前行的所有表格
            sheet.cell(row=i,column=j).border = border  #将当前行的每一个表格填充颜色
            sheet.row_dimensions[i].height = 30

            # sheet.cell(row=i,column=j).alignment = Alignment(wrap_text=True)
            sheet.cell(row=i,column=j).alignment = align
            if i ==1:
                sheet.cell(row=i,column=j).font = font1
                sheet.row_dimensions[i].height = 85
            else:
                sheet.cell(row=i, column=j).font = font2

        i=i+1 #遍历下一行


    excel.save(r'C:\Users\wb\Desktop\LXGG1.xlsx')

# write_excel()
