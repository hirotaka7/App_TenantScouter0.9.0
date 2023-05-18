import streamlit as st
import pandas as pd
import numpy as np
import folium as fl
import geopandas as gpd
from folium.plugins import BeautifyIcon
import zipfile
from PIL import Image
import os
from io import BytesIO
from xlsxwriter import Workbook
import jaconv
import math

BaseColor=['#80BBAD', '#435254', '#17E88F', '#DBD99A', '#D2785A', '#885073', '#A388BF', '#1F3765', '#3E7CA6', '#CAD1D3']
PU_Start='<div style="font-size: 10pt; color : #435254; font-weight: bold;"><span style="white-space: nowrap;">'
PU_End='</span></div>'

df_MsbGeo_dtype={
    'Comp_Name':str,
    'Branch_Name':str,
    'Branch_1':str,
    'Branch_2':str,
    'Branch_3':str,
    'Address':str,
    'Address_Type':str,
    'PostCode':int,
    'CorpNum':str,
    'Capital':float,
    'Revenue':float,
    'Num_Branch':float,
    'Num_Employee':float,
    'Num_Factory':float,
    'Num_Shop':float,
    'Ind_Main1':str,
    'Ind_Main2':str,
    'Ind_Sub1':str,
    'Ind_Sub2':str,
    'POINT_X':float,
    'POINT_Y':float
}

mdf_MsbInd=pd.read_csv('Musubu_Industry_Master.csv',encoding='utf-8')

st.set_page_config(layout="wide")
st.title('Tenant Scouter APP')

with st.sidebar:
    zip_Capsule=st.file_uploader(label='Upload Capsule.zip',type='zip')
    if zip_Capsule is not None:
        with zipfile.ZipFile(zip_Capsule,'r') as z:
            with z.open('02_PropGeo.zip') as f:
                gdf_PropGeo=gpd.read_file(f,encoding='utf-8')
            with z.open('03_Ring.zip') as f:
                gdf_Ring=gpd.read_file(f,encoding='utf-8')
            with z.open('06_MsbGeo.csv') as f:
                df_MsbGeo=pd.read_csv(f,encoding='utf-8',dtype=df_MsbGeo_dtype)
                df=pd.DataFrame(df_MsbGeo.Comp_Name.value_counts()).reset_index().rename(columns={
                    'Comp_Name':'周辺拠点数','index':'企業名'
                })

if zip_Capsule is None:
    image=Image.open('ACube.PNG')
    st.image(image,width=400)
    

if zip_Capsule is not None:       
    with st.sidebar:
        PropLabel=gdf_PropGeo.Property.values[0]
        st.header(PropLabel)
        SMap=fl.Map(
            location=[gdf_PropGeo.loc[0,'POINT_Y'],gdf_PropGeo.loc[0,'POINT_X']],
            tiles='cartodbpositron',
            zoom_start=8
        )
        Group=fl.FeatureGroup(name=gdf_PropGeo.loc[0,'Label'],show=True).add_to(SMap)
        PU='<br>'.join([gdf_PropGeo.loc[0,'Label'],gdf_PropGeo.loc[0,'Property'],gdf_PropGeo.loc[0,'Address']])
        Group.add_child(
            fl.Marker(
                location=[gdf_PropGeo.loc[0,'POINT_Y'],gdf_PropGeo.loc[0,'POINT_X']],
                popup=PU_Start+PU+PU_End,
                icon=BeautifyIcon(icon='star',border_width=2,border_color=BaseColor[0],text_color=BaseColor[0],spin=True)   
            )
        )
        st.components.v1.html(fl.Figure().add_child(SMap).render(),height=300)
        
        Kilo=gdf_Ring.kilo.values[-1]
        KiloLabel=f'{Kilo}圏内データ'
        st.header(KiloLabel)
    SearchMode=st.radio(
        label='検索モード | Search Mode',
        options=['企業フィルターサーチ | Company Filter Search','類似企業サーチ | Similar Company Search']
    )
    
    st.markdown('--------')
    
    col1,col2=st.columns(2)
    with col1:
        if SearchMode=='企業フィルターサーチ | Company Filter Search':
            Ind1_List=st.multiselect(
                label='大業界 | Large Industry',
                options=df_MsbGeo.Ind_Main1.value_counts().index,
                default=None
            )
            Ind1_MainSub=st.radio(
                label='メイン/サブ | Main/Sub',
                options=['メイン大業界のみ','サブ大業界も含む']
            )
            if len(Ind1_List)!=0:
                if Ind1_MainSub=='メイン大業界のみ':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            df_MsbGeo.Ind_Main1.isin(Ind1_List)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main1='●').rename(
                            columns={'Ind_Main1':'大業界'}
                        ),
                        how='outer'
                    )
                if Ind1_MainSub=='サブ大業界も含む':
                    df_Ind_Sub1=df_MsbGeo.Ind_Sub1.str.split(', ',expand=True)
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Ind_Main1.isin(Ind1_List))|
                            (df_MsbGeo.index.isin(df_Ind_Sub1[df_Ind_Sub1.isin(Ind1_List)].dropna(how='all').index))
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main1='●').rename(
                            columns={'Ind_Main1':'大業界'}
                        ),
                        how='outer'
                    )
            Ind2_List=st.multiselect(
                label='小業界 | Small Industry',
                options=mdf_MsbInd[mdf_MsbInd.Large_Ind.isin(Ind1_List)].Small_Ind.values,
                default=None
            )
            Ind2_MainSub=st.radio(
                label='メイン/サブ | Main/Sub',
                options=['メイン小業界のみ','サブ小業界も含む']
            )
            if len(Ind2_List)!=0:
                if Ind2_MainSub=='メイン小業界のみ':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            df_MsbGeo.Ind_Main2.isin(Ind2_List)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main2='●').rename(
                            columns={'Ind_Main2':'小業界'}
                        ),
                        how='outer'
                    )
                if Ind2_MainSub=='サブ小業界も含む':
                    df_Ind_Sub2=df_MsbGeo.Ind_Sub2.str.split(', ',expand=True)
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Ind_Main2.isin(Ind2_List))|
                            (df_MsbGeo.index.isin(df_Ind_Sub2[df_Ind_Sub2.isin(Ind2_List)].dropna(how='all').index))
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main2='●').rename(
                            columns={'Ind_Main2':'小業界'}
                        ),
                        how='outer'
                    )
            start_Capital,end_Capital=st.select_slider(
                label='資本金 | Capital',
                options=[
                    '下限なし',
                    '5,000万円',
                    '3億円',
                    '10億円',
                    '50億円',
                    '150億円',
                    '上限なし'
                ],
                value=('下限なし','上限なし')
            )
            if start_Capital=='下限なし':
                s_Capital=0
            if start_Capital=='5,000万円':
                s_Capital=50000000
            if start_Capital=='3億円':
                s_Capital=300000000
            if start_Capital=='10億円':
                s_Capital=1000000000
            if start_Capital=='50億円':
                s_Capital=5000000000
            if start_Capital=='150億円':
                s_Capital=15000000000
            if start_Capital=='上限なし':
                s_Capital=df_MsbGeo.Capital.max()
            if end_Capital=='下限なし':
                e_Capital=0
            if end_Capital=='5,000万円':
                e_Capital=50000000
            if end_Capital=='3億円':
                e_Capital=300000000
            if end_Capital=='10億円':
                e_Capital=1000000000
            if end_Capital=='50億円':
                e_Capital=5000000000
            if end_Capital=='150億円':
                e_Capital=15000000000
            if end_Capital=='上限なし':
                e_Capital=df_MsbGeo.Capital.max()
            if start_Capital!='下限なし' or end_Capital!='上限なし':
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Capital>=s_Capital)&(df_MsbGeo.Capital<e_Capital)
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Capital='●').rename(
                        columns={'Capital':'資本金'}
                    ),
                    how='outer'
                )
            start_Revenue,end_Revenue=st.select_slider(
                label='売上高 | Revenue',
                options=[
                    '下限なし',
                    '1億円',
                    '3億円',
                    '10億円',
                    '50億円',
                    '300億円',
                    '1,000億円',
                    '上限なし'
                ],
                value=('下限なし','上限なし')
            )
            if start_Revenue=='下限なし':
                s_Revenue=0
            if start_Revenue=='1億円':
                s_Revenue=100000000
            if start_Revenue=='3億円':
                s_Revenue=300000000
            if start_Revenue=='10億円':
                s_Revenue=1000000000
            if start_Revenue=='50億円':
                s_Revenue=5000000000
            if start_Revenue=='300億円':
                s_Revenue=30000000000
            if start_Revenue=='1,000億円':
                s_Revenue=100000000000
            if start_Revenue=='上限なし':
                s_Revenue=df_MsbGeo.Revenue.max()
            if end_Revenue=='下限なし':
                e_Revenue=0
            if end_Revenue=='1億円':
                e_Revenue=100000000
            if end_Revenue=='3億円':
                e_Revenue=300000000
            if end_Revenue=='10億円':
                e_Revenue=1000000000
            if end_Revenue=='50億円':
                e_Revenue=5000000000
            if end_Revenue=='300億円':
                e_Revenue=30000000000
            if end_Revenue=='1,000億円':
                e_Revenue=100000000000
            if end_Revenue=='上限なし':
                e_Revenue=df_MsbGeo.Revenue.max()
            if start_Revenue!='下限なし' or end_Revenue!='上限なし':
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Revenue>=s_Revenue)&(df_MsbGeo.Revenue<e_Revenue)
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Revenue='●').rename(
                        columns={'Revenue':'売上高'}
                    ),
                    how='outer'
                )
            start_Employee,end_Employee=st.select_slider(
                label='従業員数 | Number of Employees',
                options=[
                    '下限なし',
                    '5人',
                    '10人',
                    '20人',
                    '50人',
                    '300人',
                    '1,000人',
                    '上限なし'
                ],
                value=('下限なし','上限なし')
            )
            if start_Employee=='下限なし':
                s_Employee=0
            if start_Employee=='5人':
                s_Employee=5
            if start_Employee=='10人':
                s_Employee=10
            if start_Employee=='20人':
                s_Employee=20
            if start_Employee=='50人':
                s_Employee=50
            if start_Employee=='300人':
                s_Employee=300
            if start_Employee=='1,000人':
                s_Employee=1000
            if start_Employee=='上限なし':
                s_Employee=df_MsbGeo.Num_Employee.max()
            if end_Employee=='下限なし':
                e_Employee=0
            if end_Employee=='5人':
                e_Employee=5
            if end_Employee=='10人':
                e_Employee=10
            if end_Employee=='20人':
                e_Employee=20
            if end_Employee=='50人':
                e_Employee=50
            if end_Employee=='300人':
                e_Employee=300
            if end_Employee=='1,000人':
                e_Employee=1000
            if end_Employee=='上限なし':
                e_Employee=df_MsbGeo.Num_Employee.max()
            if start_Employee!='下限なし' or end_Employee!='上限なし':
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Num_Employee>=s_Employee)&(df_MsbGeo.Num_Employee<e_Employee)
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Num_Employee='●').rename(
                        columns={'Num_Employee':'従業員数'}
                    ),
                    how='outer'
                )
            start_Branch,end_Branch=st.select_slider(
                label='事業所数 | Number of Branches',
                options=[
                    '下限なし',
                    '1カ所',
                    '2カ所',
                    '4カ所',
                    '10カ所',
                    '20カ所',
                    '50カ所',
                    '上限なし'
                ],
                value=('下限なし','上限なし')
            )
            if start_Branch=='下限なし':
                s_Branch=0
            if start_Branch=='1カ所':
                s_Branch=1
            if start_Branch=='2カ所':
                s_Branch=2
            if start_Branch=='4カ所':
                s_Branch=4
            if start_Branch=='10カ所':
                s_Branch=10
            if start_Branch=='20カ所':
                s_Branch=20
            if start_Branch=='50カ所':
                s_Branch=50
            if start_Branch=='上限なし':
                s_Branch=df_MsbGeo.Num_Branch.max()
            if end_Branch=='下限なし':
                e_Branch=0
            if end_Branch=='1カ所':
                e_Branch=1
            if end_Branch=='2カ所':
                e_Branch=2
            if end_Branch=='4カ所':
                e_Branch=4
            if end_Branch=='10カ所':
                e_Branch=10
            if end_Branch=='20カ所':
                e_Branch=20
            if end_Branch=='50カ所':
                e_Branch=50
            if end_Branch=='上限なし':
                e_Branch=df_MsbGeo.Num_Branch.max()
            if start_Branch!='下限なし' or end_Branch!='上限なし':
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Num_Branch>=s_Branch)&(df_MsbGeo.Num_Branch<e_Branch)
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Num_Branch='●').rename(
                        columns={'Num_Branch':'事業所数'}
                    ),
                    how='outer'
                )
            df_Search=pd.DataFrame(
                data=[
                    PropLabel,
                    SearchMode,
                    Ind1_List,
                    Ind1_MainSub,
                    Ind2_List,
                    Ind2_MainSub,
                    f'{start_Capital} - {end_Capital}',
                    f'{start_Revenue} - {end_Revenue}',
                    f'{start_Employee} - {end_Employee}',
                    f'{start_Branch} - {end_Branch}',
                ],
                index=[
                    '対象物件 | Target Property',
                    '検索モード | Search Mode',
                    '大業界 | Large Industry',
                    'メイン/サブ | Main/Sub',
                    '小業界 | Small Industry',
                    'メイン/サブ | Main/Sub',
                    '資本金 | Capital',
                    '売上高 | Revenue',
                    '従業員数 | Number of Employees',
                    '事業所数 | Number of Branches',
                ],
                columns=['検索条件 | Search Condition']
            )
###############################################################################################
        if SearchMode=='類似企業サーチ | Similar Company Search':
            SearchMethod=st.radio('企業検索方法 | Company Select Method',('リスト | List','キーワード | Key Word'))
            if SearchMethod=='リスト | List':
                Comp_Name=st.selectbox(
                    '企業名 | Company Name',
                    df_MsbGeo.sort_values('Comp_Name').drop_duplicates('Comp_Name').Comp_Name.values
                )
            if SearchMethod=='キーワード | Key Word':
                Comp_KW=st.text_input(label='キーワード | Key Word : ',value='')
                if Comp_KW=='':
                    Comp_Name=''
                else:
                    Comp_KW=jaconv.h2z(Comp_KW,digit=True,ascii=True)
                    if len(df_MsbGeo[df_MsbGeo.Comp_Name.str.contains(Comp_KW)])!=0:
                        Comp_Name=st.selectbox(
                            '企業名 | Company Name',
                            df_MsbGeo[df_MsbGeo.Comp_Name.str.contains(Comp_KW)]\
                            .sort_values('Comp_Name').drop_duplicates('Comp_Name').Comp_Name.values
                        )
                    else:
                        Comp_Name=''
            if Comp_Name!='':
                CorpNum=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].CorpNum.values[0]
                Ind_Main1=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Ind_Main1.values[0]
                Ind_Main2=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Ind_Main2.values[0]
                Ind_Sub1=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Ind_Sub1.values[0]
                Ind_Sub2=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Ind_Sub2.values[0]
                Capital=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Capital.values[0]
                if math.isnan(Capital):
                    Capital_Range='-'
                else:
                    if Capital<50000000:
                        Capital_Range='[5,000万円未満]'
                        s_Capital=0
                        e_Capital=50000000
                    if 50000000<=Capital<300000000:
                        Capital_Range='[5,000万円以上~3億円未満]'
                        s_Capital=50000000
                        e_Capital=df_MsbGeo.Capital.max()
                    if 300000000<=Capital<1000000000:
                        Capital_Range='[3億円以上~10億円未満]'
                        s_Capital=300000000
                        e_Capital=df_MsbGeo.Capital.max()
                    if 1000000000<=Capital<5000000000:
                        Capital_Range='[10億円以上~50億円未満]'
                        s_Capital=1000000000
                        e_Capital=df_MsbGeo.Capital.max()
                    if 5000000000<=Capital<15000000000:
                        Capital_Range='[50億円以上~150億円未満]'
                        s_Capital=5000000000
                        e_Capital=df_MsbGeo.Capital.max()
                    if 15000000000<=Capital:
                        Capital_Range='[150億円以上]'
                        s_Capital=15000000000
                        e_Capital=df_MsbGeo.Capital.max()
                try:
                    Capital='{:,}'.format(int(Capital))
                except ValueError:
                    Capital='-'
                Revenue=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Revenue.values[0]
                if math.isnan(Revenue):
                    Revenue_Range='-'
                else:
                    if Revenue<100000000:
                        Revenue_Range='[1億円未満]'
                        s_Revenue=0
                        e_Revenue=100000000
                    if 100000000<=Revenue<300000000:
                        Revenue_Range='[1億円以上~3億円未満]'
                        s_Revenue=100000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                    if 300000000<=Revenue<1000000000:
                        Revenue_Range='[3億円以上~10億円未満]'
                        s_Revenue=300000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                    if 1000000000<=Revenue<5000000000:
                        Revenue_Range='[10億円以上~50億円未満]'
                        s_Revenue=1000000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                    if 5000000000<=Revenue<30000000000:
                        Revenue_Range='[50億円以上~300億円未満]'
                        s_Revenue=5000000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                    if 30000000000<=Revenue<100000000000:
                        Revenue_Range='[300億円以上~1,000億円未満]'
                        s_Revenue=30000000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                    if 100000000000<=Revenue:
                        Revenue_Range='[1,000億円以上]'
                        s_Revenue=100000000000
                        e_Revenue=df_MsbGeo.Revenue.max()
                try:
                    Revenue='{:,}'.format(int(Revenue))
                except ValueError:
                    Revenue='-'
                Num_Employee=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Num_Employee.values[0]
                if math.isnan(Num_Employee):
                    Num_Employee_Range='-'
                else:
                    if Num_Employee<5:
                        Num_Employee_Range='[5人未満]'
                        s_Employee=0
                        e_Employee=5
                    if 5<=Num_Employee<10:
                        Num_Employee_Range='[5人以上~10人未満]'
                        s_Employee=5
                        e_Employee=df_MsbGeo.Num_Employee.max()
                    if 10<=Num_Employee<20:
                        Num_Employee_Range='[10人以上~20人未満]'
                        s_Employee=10
                        e_Employee=df_MsbGeo.Num_Employee.max()
                    if 20<=Num_Employee<50:
                        Num_Employee_Range='[20人以上~50人未満]'
                        s_Employee=20
                        e_Employee=df_MsbGeo.Num_Employee.max()
                    if 50<=Num_Employee<300:
                        Num_Employee_Range='[50人以上~300人未満]'
                        s_Employee=50
                        e_Employee=df_MsbGeo.Num_Employee.max()
                    if 300<=Num_Employee<1000:
                        Num_Employee_Range='[300人以上~1000人未満]'
                        s_Employee=300
                        e_Employee=df_MsbGeo.Num_Employee.max()
                    if 1000<=Num_Employee:
                        Num_Employee_Range='[1000人以上]'
                        s_Employee=100
                        e_Employee=df_MsbGeo.Num_Employee.max()
                try:
                    Num_Employee='{:,}'.format(int(Num_Employee))
                except ValueError:
                    Num_Employee='-'
                Num_Branch=df_MsbGeo[df_MsbGeo.Comp_Name==Comp_Name].Num_Branch.values[0]
                if math.isnan(Num_Branch):
                    Num_Branch_Range='-'
                else:
                    if Num_Branch<2:
                        Num_Branch_Range='[1カ所以上~2カ所未満]'
                        s_Branch=0
                        e_Branch=2
                    if 2<=Num_Branch<4:
                        Num_Branch_Range='[2カ所以上~4カ所未満]'
                        s_Branch=2
                        e_Branch=df_MsbGeo.Num_Branch.max()
                    if 4<=Num_Branch<10:
                        Num_Branch_Range='[4カ所以上~10カ所未満]'
                        s_Branch=4
                        e_Branch=df_MsbGeo.Num_Branch.max()
                    if 10<=Num_Branch<20:
                        Num_Branch_Range='[10カ所以上~20カ所未満]'
                        s_Branch=10
                        e_Branch=df_MsbGeo.Num_Branch.max()
                    if 20<=Num_Branch<50:
                        Num_Branch_Range='[20カ所以上~50カ所未満]'
                        s_Branch=20
                        e_Branch=df_MsbGeo.Num_Branch.max()
                    if 50<=Num_Branch:
                        Num_Branch_Range='[50カ所以上]'
                        s_Branch=50
                        e_Branch=df_MsbGeo.Num_Branch.max()
                try:
                    Num_Branch='{:,}'.format(int(Num_Branch))
                except ValueError:
                    Num_Branch='-'
                
            else:
                CorpNum='-'
                Ind_Main1='-'
                Ind_Main2='-'
                Ind_Sub1='-'
                Ind_Sub2='-'
                Capital='-'
                Capital_Range='-'
                Revenue='-'
                Revenue_Range='-'
                Num_Employee='-'
                Num_Employee_Range='-'
                Num_Branch='-'
                Num_Branch_Range='-'
            st.caption('企業名 | Company Name')
            st.subheader(Comp_Name)
            scol1,scol2=st.columns(2)
            with scol1:    
                st.caption('法人番号 | Corporate Number')
                st.text(CorpNum)
                st.caption('メイン大業界 | Main Large Industry')
                st.text(Ind_Main1)
                st.caption('メイン小業界 | Main Small Industry')
                st.text(Ind_Main2)
                st.caption('サブ大業界 | Sub Large Industry')
                st.text(Ind_Sub1.replace(', ','\n'))
                st.caption('サブ小業界 | Sub Small Industry')
                st.text(Ind_Sub2.replace(', ','\n'))
            with scol2:
                st.caption('資本金 | Capital')
                st.text(Capital+'\n'+Capital_Range)
                st.caption('売上高 | Revenue')
                st.text(Revenue+'\n'+Revenue_Range)
                st.caption('従業員数 | Number of Employees')
                st.text(Num_Employee+'\n'+Num_Employee_Range)
                st.caption('事業所数 | Number of Branches')
                st.text(Num_Branch+'\n'+Num_Branch_Range)
            if CorpNum!='-':
                df_Ind_Sub1=df_MsbGeo.Ind_Sub1.str.split(', ',expand=True)
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Ind_Main1==Ind_Main1)|
                        (df_MsbGeo.index.isin(df_Ind_Sub1[df_Ind_Sub1.isin([Ind_Main1])].dropna(how='all').index))
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main1='●').rename(
                        columns={'Ind_Main1':'大業界'}
                    ),
                    how='outer'
                )
                df_Ind_Sub2=df_MsbGeo.Ind_Sub2.str.split(', ',expand=True)
                df=pd.merge(
                    df,
                    df_MsbGeo[
                        (df_MsbGeo.Ind_Main2==Ind_Main2)|
                        (df_MsbGeo.index.isin(df_Ind_Sub2[df_Ind_Sub2.isin([Ind_Main2])].dropna(how='all').index))
                    ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Ind_Main2='●').rename(
                        columns={'Ind_Main2':'小業界'}
                    ),
                    how='outer'
                )
                if Capital_Range!='-':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Capital>=s_Capital)&(df_MsbGeo.Capital<e_Capital)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Capital='●').rename(
                            columns={'Capital':'資本金'}
                        ),
                        how='outer'
                    )
                if Revenue_Range!='-':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Revenue>=s_Revenue)&(df_MsbGeo.Revenue<e_Revenue)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Revenue='●').rename(
                            columns={'Revenue':'売上高'}
                        ),
                        how='outer'
                    )
                if Num_Employee_Range!='-':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Num_Employee>=s_Employee)&(df_MsbGeo.Num_Employee<e_Employee)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Num_Employee='●').rename(
                            columns={'Num_Employee':'従業員数'}
                        ),
                        how='outer'
                    )
                if Num_Branch_Range!='-':
                    df=pd.merge(
                        df,
                        df_MsbGeo[
                            (df_MsbGeo.Num_Branch>=s_Branch)&(df_MsbGeo.Num_Branch<e_Branch)
                        ][['Comp_Name']].drop_duplicates().rename(columns={'Comp_Name':'企業名'}).assign(Num_Branch='●').rename(
                            columns={'Num_Branch':'事業所数'}
                        ),
                        how='outer'
                    )
                df_Search=pd.DataFrame(
                    data=[
                        PropLabel,
                        SearchMode,
                        Comp_Name,
                        Ind_Main1,
                        Ind_Main2,
                        Capital_Range,
                        Revenue_Range,
                        Num_Employee_Range,
                        Num_Branch_Range,

                    ],
                    index=[
                        '対象物件 | Target Property',
                        '検索モード | Search Mode',
                        '企業名 | Company Name',
                        'メイン大業界 | Main Large Industry',
                        'メイン小業界 | Main Small Industry',
                        '資本金 | Capital',
                        '売上高 | Revenue',
                        '従業員数 | Number of Employees',
                        '事業所数 | Number of Branches'
                    ],
                    columns=['検索条件 | Search Condition']
                )
###############################################################################################
    
    with col2:
        df=df.assign(Point=(df=='●').sum(axis=1)).sort_values(['Point','周辺拠点数'],ascending=False)
        ColList=df.columns.tolist()
        NewColList=ColList[0:1]+ColList[-1:]+ColList[1:-1]
        df=df[NewColList].reset_index(drop=True)
        st.dataframe(df,height=750,use_container_width=True)
    with st.sidebar:
        st.metric('該当企業数 | Number of Companies',len(df.dropna()))
        def df_to_xlsx(df_Search,df):
            byte_xlsx=BytesIO()
            writer_xlsx=pd.ExcelWriter(byte_xlsx, engine="xlsxwriter")
            df_Search.to_excel(writer_xlsx,index=True,sheet_name="検索条件 | Search Condition")
            df.to_excel(writer_xlsx,index=False,sheet_name="検索結果 | Search Result")
            writer_xlsx.close()
            out_xlsx=byte_xlsx.getvalue()
            return out_xlsx
        out_xlsx=df_to_xlsx(df_Search,df)
        st.download_button('📋 Download Excel',out_xlsx,file_name='TenantScouter.xlsx')
        


