import shutil
import re
import time
import streamlit as st
import pandas as pd
import os
import openpyxl
from datetime import datetime

# 清理非法字符的函数

# @st.cache
def clean_illegal_characters(value):
    """
    清理非法字符
    """
    if isinstance(value, str):
        # 移除非法字符
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value)
    return value

# 合并 CSV 文件的函数
def merge_csv_to_excel(folder_path):
    """
    将指定文件夹中的所有 CSV 文件合并为一个 Excel 文件，删除重复表头，保存到指定路径。
    """
    # 初始化一个空的DataFrame用于存储合并的数据
    combined_df = pd.DataFrame()
    
    # 获取文件夹中的所有CSV文件
    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    
    # 展示进度条
    progress_bar = st.progress(0)
    for i, csv_file in enumerate(csv_files):
        csv_file_path = os.path.join(folder_path, csv_file)
        st.write(f"正在处理文件: {csv_file_path}")
        
        # 读取CSV文件
        df = pd.read_csv(csv_file_path)
        
        # 清理数据中的非法字符
        df = df.applymap(clean_illegal_characters)
        
        # 如果combined_df为空，则直接赋值
        if combined_df.empty:
            combined_df = df
        else:
            # 否则，跳过当前CSV的表头行并追加数据
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        
        # 更新进度条
        progress_bar.progress((i + 1) / len(csv_files))
        time.sleep(0.1)  # 模拟处理延迟
    
    # 将合并后的数据写入Excel文件
    # combined_df.to_excel(output_excel_path, index=False, engine='openpyxl')
    # st.success(f"合并后的文件已保存为: {output_excel_path}")
    combined_df.columns = combined_df.columns.str.lower()
    st.write(f"合并后的列名: {combined_df.columns.tolist()}")  # 打印列名
    
    # 返回合并后的 DataFrame
    return combined_df

# 转换 Excel 数据的函数
def transform_data(df, output_file: str):
    """
    读取合并后的 Excel 文件并创建一个新的 Excel 文件，根据指定规则转换数据。
    """
    # 国家代码到名称的映射
    code_to_country_mapping = {
        "af": "Afghanistan", "ax": "Åland", "al": "Albania", "dz": "Algeria",
        "as": "American Samoa", "ad": "Andorra", "ao": "Angola", "ai": "Anguilla",
        "aq": "Antarctica", "ag": "Antigua and Barbuda", "ar": "Argentina",
        "am": "Armenia", "aw": "Aruba", "ac": "Ascension Island", "au": "Australia",
        "at": "Austria", "az": "Azerbaijan", "bs": "Bahamas", "bh": "Bahrain",
        "bd": "Bangladesh", "bb": "Barbados", "eus": "Basque Country", "by": "Belarus",
        "be": "Belgium", "bz": "Belize", "bj": "Benin", "bm": "Bermuda", "bt": "Bhutan",
        "bo": "Bolivia", "bq": "Bonaire, Saba, and Sint Eustatius", "ba": "Bosnia and Herzegovina",
        "bw": "Botswana", "bv": "Bouvet Island", "br": "Brazil", "io": "British Indian Ocean Territory",
        "vg": "British Virgin Islands", "bn": "Brunei", "bg": "Bulgaria", "bf": "Burkina Faso",
        "mm": "Burma (officially: Myanmar)", "bi": "Burundi", "kh": "Cambodia", "cm": "Cameroon",
        "ca": "Canada", "cv": "Cape Verde", "cat": "Catalonia", "ky": "Cayman Islands",
        "cf": "Central African Republic", "td": "Chad", "cl": "Chile", "cn": "China, People’s Republic of",
        "cx": "Christmas Island", "cc": "Cocos (Keeling) Islands", "co": "Colombia", "km": "Comoros",
        "cd": "Congo, Democratic Republic of the (Congo-Kinshasa)", "cg": "Congo, Republic of the (Congo-Brazzaville)",
        "ck": "Cook Islands", "cr": "Costa Rica", "ci": "Côte d’Ivoire (Ivory Coast)", "hr": "Croatia",
        "cu": "Cuba", "cw": "Curaçao", "cy": "Cyprus", "cz": "Czech Republic", "dk": "Denmark",
        "dj": "Djibouti", "dm": "Dominica", "do": "Dominican Republic", "tl": "East Timor (Timor-Leste)",
        "ec": "Ecuador", "eg": "Egypt", "sv": "El Salvador", "gq": "Equatorial Guinea", "er": "Eritrea",
        "ee": "Estonia", "et": "Ethiopia", "eu": "European Union", "fk": "Falkland Islands",
        "fo": "Faeroe Islands", "fm": "Federated States of Micronesia", "fj": "Fiji", "fi": "Finland",
        "fr": "France", "gf": "French Guiana", "pf": "French Polynesia", "tf": "French Southern and Antarctic Lands",
        "ga": "Gabon (officially: Gabonese Republic)", "gal": "Galicia", "gm": "Gambia", "ps": "Gaza Strip (Gaza)",
        "ge": "Georgia", "de": "Germany", "gh": "Ghana", "gi": "Gibraltar", "gr": "Greece", "gl": "Greenland",
        "gd": "Grenada", "gp": "Guadeloupe", "gu": "Guam", "gt": "Guatemala", "gg": "Guernsey", "gn": "Guinea",
        "gw": "Guinea-Bissau", "gy": "Guyana", "ht": "Haiti", "hm": "Heard Island and McDonald Islands",
        "hn": "Honduras", "hk": "Hong Kong", "hu": "Hungary", "is": "Iceland", "in": "India", "id": "Indonesia",
        "ir": "Iran", "iq": "Iraq", "ie": "Ireland", "im": "Isle of Man", "il": "Israel", "it": "Italy",
        "jm": "Jamaica", "jp": "Japan", "je": "Jersey", "jo": "Jordan", "kz": "Kazakhstan", "ke": "Kenya",
        "ki": "Kiribati", "xk": "Kosovo", "kw": "Kuwait", "kg": "Kyrgyzstan", "la": "Laos", "lv": "Latvia",
        "lb": "Lebanon", "ls": "Lesotho", "lr": "Liberia", "ly": "Libya", "li": "Liechtenstein", "lt": "Lithuania",
        "lu": "Luxembourg", "mo": "Macau", "mk": "Macedonia, Republic of (the former Yugoslav Republic of Macedonia, FYROM)",
        "mg": "Madagascar", "mw": "Malawi", "my": "Malaysia", "mv": "Maldives", "ml": "Mali", "mt": "Malta",
        "mh": "Marshall Islands", "mq": "Martinique", "mr": "Mauritania", "mu": "Mauritius", "yt": "Mayotte",
        "mx": "Mexico", "md": "Moldova", "mc": "Monaco", "mn": "Mongolia", "me": "Montenegro", "ms": "Montserrat",
        "ma": "Morocco", "mz": "Mozambique", "mm": "Myanmar", "na": "Namibia", "nr": "Nauru", "np": "Nepal",
        "nl": "Netherlands", "nc": "New Caledonia", "nz": "New Zealand", "ni": "Nicaragua", "ne": "Niger",
        "ng": "Nigeria", "nu": "Niue", "nf": "Norfolk Island", "nctr": "North Cyprus (unrecognised, self-declared state)",
        "kp": "North Korea", "mp": "Northern Mariana Islands", "no": "Norway", "om": "Oman", "pk": "Pakistan",
        "pw": "Palau", "ps": "Palestine", "pa": "Panama", "pg": "Papua New Guinea", "py": "Paraguay", "pe": "Peru",
        "ph": "Philippines", "pn": "Pitcairn Islands", "pl": "Poland", "pt": "Portugal", "pr": "Puerto Rico",
        "qa": "Qatar", "ro": "Romania", "ru": "Russia", "rw": "Rwanda", "re": "Réunion Island",
        "bl": "Saint Barthélemy (informally also referred to as Saint Barth’s or Saint Barts)", "sh": "Saint Helena",
        "kn": "Saint Kitts and Nevis", "lc": "Saint Lucia", "mf": "Saint Martin (officially the Collectivity of Saint Martin)",
        "pm": "Saint-Pierre and Miquelon", "vc": "Saint Vincent and the Grenadines", "ws": "Samoa", "sm": "San Marino",
        "st": "São Tomé and Príncipe", "sa": "Saudi Arabia", "sn": "Senegal", "rs": "Serbia", "sc": "Seychelles",
        "sl": "Sierra Leone", "sg": "Singapore", "an": "Sint Maarten", "sk": "Slovakia", "si": "Slovenia",
        "sb": "Solomon Islands", "so": "Somalia", "so": "Somaliland", "za": "South Africa",
        "gs": "South Georgia and the South Sandwich Islands", "kr": "South Korea", "ss": "South Sudan", "es": "Spain",
        "lk": "Sri Lanka", "sd": "Sudan", "sr": "Suriname", "sj": "Svalbard and Jan Mayen Islands", "sz": "Swaziland",
        "se": "Sweden", "ch": "Switzerland", "sy": "Syria", "tw": "Taiwan", "tj": "Tajikistan", "tz": "Tanzania",
        "th": "Thailand", "tg": "Togo", "tk": "Tokelau", "to": "Tonga", "tt": "Trinidad & Tobago", "tn": "Tunisia",
        "tr": "Turkey", "tm": "Turkmenistan", "tc": "Turks and Caicos Islands", "tv": "Tuvalu", "ug": "Uganda",
        "ua": "Ukraine", "ae": "United Arab Emirates (UAE)", "gb": "United Kingdom (UK)", "us": "United States of America (USA)",
        "vi": "United States Virgin Islands", "xx": "Unknown", "uy": "Uruguay", "uz": "Uzbekistan", "vu": "Vanuatu",
        "va": "Vatican City", "ve": "Venezuela", "vn": "Vietnam", "wf": "Wallis and Futuna", "eh": "Western Sahara",
        "ye": "Yemen", "zm": "Zambia", "zw": "Zimbabwe"
    }

    
    # 创建输出文件
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    
    # 指定新的表头
    new_header = [
        "Platform", "Date", "Time", "Text", "Permalink", "Country", "Engagement",
        "comments count", "likes count", "shares count", "Author Name", "Screen Name",
        "120Hz", "5000 mAh","50MP","AI","IP67","32MP","Vivid Nightography","Fast charging","HDR","OIS","Super AMOLED","Samsung Knox Vault","VDIS"
    ]
    
    # 写入新的表头
    new_sheet.append(new_header)
    
    # 获取总行数
    total_rows = len(df)
    progress_bar = st.progress(0)  # 进度条
    processed_rows = 0  # 已处理的行数
    
    # 复制并转换数据
    for index, row in df.iterrows():
        platform = row["platform"]
        date_time_str = row["date"]
        text = row["text"]
        permalink = row["permalink"]
        key_markets = row["key markets"]
        inferred_country = row["inferred country"]
        author = row["author"]
        screen_name = row["screen name"]
        engagement_actions = row["engagement actions"]
        samsung_a55_ksp = row["samsung a55 ksp"]
        
     # 创建输出文件
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    
    # 指定新的表头
    new_header = [
        "Platform", "Date", "Time", "Text", "Permalink", "Country", "Engagement",
        "comments count", "likes count", "shares count", "Author Name", "Screen Name",
        "120Hz", "5000 mAh","50MP","AI","IP67","32MP","Vivid Nightography","Fast charging","HDR","OIS","Super AMOLED","Samsung Knox Vault","VDIS"
    ]
    
    # 写入新的表头
    new_sheet.append(new_header)
    
    # 获取总行数
    total_rows = len(df)
    progress_bar = st.progress(0)  # 进度条
    processed_rows = 0  # 已处理的行数
    
    # 复制并转换数据
    for index, row in df.iterrows():
        platform = str(row["platform"]) if pd.notna(row["platform"]) else ""
        date_time_str = str(row["date"]) if pd.notna(row["date"]) else ""
        text = str(row["text"]) if pd.notna(row["text"]) else ""
        permalink = str(row["permalink"]) if pd.notna(row["permalink"]) else ""
        key_markets = str(row["key markets"]) if pd.notna(row["key markets"]) else ""
        inferred_country = str(row["inferred country"]) if pd.notna(row["inferred country"]) else ""
        author = str(row["author"]) if pd.notna(row["author"]) else ""
        screen_name = str(row["screen name"]) if pd.notna(row["screen name"]) else ""
        engagement_actions = str(row["engagement actions"]) if pd.notna(row["engagement actions"]) else ""
        samsung_a55_ksp = str(row["samsung a55 ksp"]) if pd.notna(row["samsung a55 ksp"]) else ""
        
        # 处理日期和时间
        if date_time_str:
            try:
                date_time_obj = datetime.strptime(date_time_str, '%d/%m/%Y %H:%M:%S')
                date = date_time_obj.date()
                time_value = date_time_obj.strftime('%H:%M:%S')  # 转换为字符串
            except ValueError:
                date = None
                time_value = None
        else:
            date = None
            time_value = None
        
        # 处理国家信息
        if "global" in key_markets.lower():
            if key_markets.lower() == "global":
                country = inferred_country
            else:
                country = key_markets.replace("global", "").strip()
        else:
            country = key_markets
        
        # 替换国家代码为名称
        country_name = code_to_country_mapping.get(country.lower(), country)
        
        # 去掉"Global"及其最近的一个分号
        if "global" in country_name.lower():
            country_name = country_name.replace("Global", "").strip()
            if ";" in country_name:
                if country_name.startswith(";"):
                    country_name = country_name[1:].strip()
                elif country_name.endswith(";"):
                    country_name = country_name[:-1].strip()
        
        # 处理参与度和社交媒体指标
        if platform.lower() == "x":
            likes_count = row["x likes"]
            comments_count = row["x replies"]
            shares_count = row["x reposts"]
            engagement = ''
        elif platform.lower() == "facebook":
            likes_count = row["facebook likes"]
            comments_count = row["facebook comments"]
            shares_count = row["facebook shares"]
            engagement = ''
        elif platform.lower() in ["blog", "reddit", "forum", "linkedin"]:
            engagement = engagement_actions
            likes_count = ''
            comments_count = ''
            shares_count = ''
        else:
            engagement = ''
            likes_count = ''
            comments_count = ''
            shares_count = ''
        
        # 处理Samsung A55 KSP列
        samsung_features = {
            "120Hz": 0,
            "5000 mAh": 0,
            "50MP": 0,
            "AI": 0,
            "IP67": 0,
            "32MP": 0,
            "Vivid Nightography": 0,
            "Fast charging": 0,
            "HDR":0,
            "OIS":0,
            "Super AMOLED":0,
            "Samsung Knox Vault":0,
            "VDIS":0
        }
        
        if samsung_a55_ksp:
            for feature in samsung_features.keys():
                if feature.lower() in samsung_a55_ksp.lower():  # 现在可以安全调用 .lower()
                    samsung_features[feature] = 1
        
        # 创建新行
        new_row = [
            platform, date, time_value, text, permalink, country_name, engagement,
            comments_count, likes_count, shares_count, author, screen_name, 
            samsung_features["120Hz"], samsung_features["5000 mAh"],
            samsung_features["50MP"], samsung_features["AI"],
            samsung_features["IP67"], samsung_features["32MP"],
            samsung_features["Vivid Nightography"], samsung_features["Fast charging"],
            samsung_features["HDR"], samsung_features["OIS"],
            samsung_features["Super AMOLED"], samsung_features["Samsung Knox Vault"],
            samsung_features["VDIS"]
        ]
        
        new_sheet.append(new_row)
        
        # 更新进度条
        processed_rows += 1
        progress = processed_rows / total_rows
        progress_bar.progress(progress)
        # st.write(f"已处理 {processed_rows}/{total_rows} 行数据...")
        time.sleep(0.1)  # 使用 time 模块的 sleep 函数
    
    # 保存输出文件
    new_wb.save(output_file)
    st.success(f"转换后的文件已保存为: {output_file}")

# Streamlit 应用界面
def main():
    st.title("CSV 文件合并与转换工具")
    st.write("上传多个 CSV 文件，将它们合并并转换为指定格式的 Excel 文件。")

    # 初始化状态变量
    if "merged" not in st.session_state:
        st.session_state.merged = False
    if "combined_df" not in st.session_state:
        st.session_state.combined_df = None

    # 文件上传组件（支持多文件上传）
    uploaded_files = st.file_uploader("上传 CSV 文件", type=["csv"], accept_multiple_files=True)

    if uploaded_files:
        # 展示已上传的文件
        st.write("### 已上传的文件：")
        for uploaded_file in uploaded_files:
            st.write(f"- {uploaded_file.name}")

        # 创建一个临时文件夹来存储上传的文件
        temp_folder = "temp_csv_files"
        
        # 清空临时文件夹
        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)
        os.makedirs(temp_folder, exist_ok=True)

        # 将上传的文件保存到临时文件夹
        for uploaded_file in uploaded_files:
            with open(os.path.join(temp_folder, uploaded_file.name), "wb") as f:
                f.write(uploaded_file.getbuffer())
        st.success("文件上传成功！")

        # 合并按钮
        if st.button("合并 CSV 文件"):
            with st.spinner("正在合并文件，请稍候..."):
                try:
                    # 调用合并函数
                    combined_df = merge_csv_to_excel(temp_folder)
                    st.session_state.merged = True  # 设置状态为已合并
                    st.session_state.combined_df = combined_df  # 保存合并后的 DataFrame
                    
                    # 展示合并后的数据预览
                    # with st.expander("点击查看合并后的数据预览", expanded=False):
                    #     st.write("### 合并后的数据预览：")
                    #     st.dataframe(combined_df)
                except Exception as e:
                    st.error(f"合并文件时出错: {e}")

        # 如果文件已合并，显示转换按钮
        if st.session_state.merged:
            output_file_path = st.text_input("请输入转换后的文件保存路径（例如：C:/Users/YourName/Desktop/转换结果.xlsx）")
            
            if output_file_path:
                if st.button("转换 Excel 文件"):
                    with st.spinner("正在转换文件，请稍候..."):
                        try:
                            # 调用转换函数
                            transform_data(st.session_state.combined_df, output_file_path)
                            
                            # 展示转换后的数据预览
                            # with st.expander("点击查看转换后的数据预览", expanded=False):
                            #     st.write("### 转换后的数据预览：")
                            #     df_transformed = pd.read_excel(output_file_path)
                            #     st.dataframe(df_transformed)
                            
                            # 提供下载链接
                            with open(output_file_path, "rb") as f:
                                st.download_button(
                                    label="下载转换后的 Excel 文件",
                                    data=f,
                                    file_name=os.path.basename(output_file_path),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        except Exception as e:
                            st.error(f"转换文件时出错: {e}")
    else:
        st.info("请上传一个或多个 CSV 文件。")

# 运行 Streamlit 应用
if __name__ == "__main__":
    main()
