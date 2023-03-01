import pandas as pd
import json

import requests


def buhuobiao(store):
    store_data = requests.get("http://43.142.117.35/get_store_id")
    store_data = store_data.json()
    df_store = pd.DataFrame()
    for i in store_data:
        df_tmp = pd.DataFrame(i,index=[0])
        df_store = df_store.append(df_tmp)
    df_store = df_store[df_store['name']==store].reset_index()
    sid = df_store.loc[0,'sid']
    offset = 0
    df_performance7D = pd.DataFrame()
    while True:
        req = {"sid": sid, 'start_date': '2023-02-05', 'end_date': '2023-02-11','summary_field':'msku','offset': offset}
        performance_data = requests.post("http://43.142.117.35/productPerFormance",params=req).json()
        performance_data = performance_data.get("data").get("list")
        if performance_data == "":
            break
        for i in performance_data:
            data = pd.DataFrame(i,index=[0])
            df_performance7D = df_performance7D.append(data)



        offset = offset + 1


    a=11


# buhuobiao("LXGG-US")

