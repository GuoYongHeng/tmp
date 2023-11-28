
import os
import re
import pdfplumber

import pandas as pd

def to_df(file_path : str, df : pd.DataFrame, cnt : int):
    with pdfplumber.open(file_path) as pdf:
        
        page = pdf.pages[ 0 ]
        text = page.extract_text()

        公司名称 = re.findall(r"购\s名\s*称[：:](\S*)", text)
        if len(公司名称) == 0:
            公司名称 = re.findall(r"\n\s*名\s*称[：:](.*)", text)
        公司名称 = 公司名称[0].strip()

        供应商名称 = re.findall(r"销\s名\s*称[：:](\S*)", text)
        if len(供应商名称) == 0:
            供应商名称 = re.findall(r"\n\s*名\s*称[：:](.*)", text)
        if len(供应商名称) == 1:
            供应商名称 = 供应商名称[0].strip()
        elif len(供应商名称) == 2:
            供应商名称 = 供应商名称[1].strip()

        发票号码 = re.findall(r"发票号码[：:]\s*(\d*)", text)[0].strip()

        发票金额 = re.findall("[¥￥](.*)[¥￥]", text)[0].strip()
        税额 = re.findall("[¥￥].*[¥￥](.*)", text)[0].strip()
        含税价格 = re.findall("价税合计.*[¥￥](.*)", text)[0].strip()

        df.loc[ cnt, "公司名称" ] = 公司名称
        df.loc[ cnt, "供应商名称" ] = 供应商名称
        df.loc[ cnt, "发票号码" ] = 发票号码
        df.loc[ cnt, "发票金额" ] = 发票金额
        df.loc[ cnt, "税额" ] = 税额
        df.loc[ cnt, "含税价格" ] = 含税价格


def init_df():
    cols = [ "序号", "Barcode Number", "公司代码", 
             "公司名称", "供应商号码", "供应商名称", 
             "供应商发票号码", "发票金额", "税额",
             "含税价格"]
    df = pd.DataFrame(columns = cols)
    return df



def main():
    path = input("请输入文件夹路径或文件路径：")
    df = init_df()
    cnt = 0
    if os.path.isdir(path):
        filenames = sorted(os.listdir(path))
        for filename in filenames:
            file_path = path + "\\" + filename
            try:
                to_df(file_path, df, cnt)
                cnt += 1
            except Exception as e:
                print(file_path, "出错")
    else:
        to_df(path, df, cnt)
    
    df.to_excel("./result.xlsx")
    

if __name__ == "__main__":
    main()