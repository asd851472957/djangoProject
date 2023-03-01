# coding=utf-8
import pandas as pd
def color_size_analysis():
    with pd.ExcelWriter(r"C:\Users\wb\Desktop\LXQW-AL-US补货-2023-01-30.xlsx", mode="a") as xlsx:
        df = pd.read_excel(r"C:\Users\wb\Desktop\LXQW-AL-US补货-2023-01-30.xlsx",sheet_name="库存明细",header=1)[['父体','IN-STOCK库存','颜色大类','细分颜色','组合方式','7天转化率']]
        df = df[df['IN-STOCK库存']>0]
        df_color_analysis = df[['父体','颜色大类','7天转化率']]
        df_color_analysis2 = df[['父体','细分颜色','7天转化率']]
        df_size_analysis = df[['父体','组合方式','7天转化率']]
        df_color_analysis = df_color_analysis.pivot_table(index=["父体","颜色大类"],values="7天转化率",aggfunc="mean").fillna(0).reset_index()
        df_color_analysis2 = df_color_analysis2.pivot_table(index=["父体","细分颜色"],values="7天转化率",aggfunc="mean").fillna(0).reset_index()
        df_size_analysis = df_size_analysis.pivot_table(index=["父体","组合方式"],values="7天转化率",aggfunc="mean").fillna(0).reset_index()
        df_color_analysis["7天转化率百分比"] = df_color_analysis["7天转化率"].map(lambda x: format(x,'.2%'))
        df_color_analysis2["7天转化率百分比"] = df_color_analysis2["7天转化率"].map(lambda x: format(x,'.2%'))
        df_size_analysis["7天转化率百分比"] = df_size_analysis["7天转化率"].map(lambda x: format(x,'.2%'))
        df_color_analysis.to_excel(xlsx,sheet_name="颜色大类分析")
        df_color_analysis2.to_excel(xlsx,sheet_name="细分颜色分析")
        df_size_analysis.to_excel(xlsx,sheet_name="尺寸分析")
        a=11


# color_size_analysis()


def color_size_analysis_all():
    # with pd.ExcelWriter(r"C:\Users\wb\Desktop\LXQW-AL-US补货-2023-01-30.xlsx", mode="a") as xlsx:
        df = pd.read_excel(r"C:\Users\wb\Desktop\LXQW-AL-US补货-2023-01-30.xlsx",sheet_name="库存明细",header=1)[['父体','IN-STOCK库存','颜色大类','细分颜色','组合方式','7天转化率']]
        df = df[df['IN-STOCK库存']>0]
        df_color_analysis = df[['父体','颜色大类','组合方式','7天转化率']]
        df_color_analysis['组合方式'] = df_color_analysis['组合方式'].astype("str")
        df_color_analysis = df_color_analysis.pivot_table(index=["父体","颜色大类","组合方式"],values="7天转化率",aggfunc="mean").fillna(0).reset_index()
        df_color_analysis["7天转化率百分比"] = df_color_analysis["7天转化率"].map(lambda x: format(x,'.2%'))
        df_color_analysis["颜色尺寸"] = df_color_analysis["颜色大类"]+df_color_analysis["组合方式"]
        df_color_analysis.to_excel(r"C:\Users\wb\Desktop\111111.xlsx",sheet_name="颜色尺寸分析")

        a=11


# color_size_analysis_all()