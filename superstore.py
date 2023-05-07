#!/usr/bin/env python
# coding: utf-8

# # 超市銷售分析
# #### 資料來源: Tableau , https://community.tableau.com/s/question/0D54T00000CWeX8SAL/sample-superstore-sales-excelxls

# In[1]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import datetime
import statistics
import warnings
warnings.filterwarnings( "ignore" )

import calendar
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.figure_factory as ff
import plotly.offline as offline
import plotly.graph_objs as go
offline.init_notebook_mode(connected = True)


# ## 一、載入資料

# In[2]:


df = pd.read_excel("https://public.tableau.com/app/sample-data/sample_-_superstore.xls")
df["Order Date"] = pd.to_datetime(df["Order Date"])


# In[3]:


df.info()


# ## 二、資料預處理
# ### 1.新增欄位 "new_customer" ，用以判斷顧客為新客或是常客，若訂單為此顧客的第一筆消費則標註為 1(新客) ，若否則標註為0(常客)
# ### 2.將欄位"Order Date"中的年、月、星期幾分別轉化出三個新欄位"year","month","day_of_week"
# ### 3.新增欄位"discount_ornot"，用以判斷此筆訂單有折扣或是原價購買，若訂單中"Discount"欄位中的值不為0則標註為1(有折扣)，若否則標註為0(無折扣)
# 

# In[4]:


df = df.sort_values(by=["Customer ID","Order Date"]).reset_index()

df_test = df
rep_num = list(df.groupby("Customer ID").size())
new_customer = []
for i in range(len(rep_num)):
    t=0
    new_customer.append(1)
    for k in range(rep_num[i]-1):
        if df_test.iloc[t,2] == df_test.iloc[k+1,2]:
            new_customer.append(1)
        else:
            new_customer.append(0)
    df_test = df_test.reset_index(drop=True).drop(range(rep_num[i]))

df.insert(4, column="new_customer", value=new_customer)


df["year"] = df["Order Date"].dt.year
df["month"] = df["Order Date"].dt.month
df["day_of_week"] = df["Order Date"].dt.day_name()


discount_ornot = []
i=0
for i in range(len(df)):
    if df["Discount"][i] == 0:
        discount_ornot.append(0)
    else:
        discount_ornot.append(1)

df.insert(18, column="discount_ornot", value=discount_ornot)


# In[5]:


df.head()


# ## 三、視覺化及分析

# In[6]:


df_y_sa = df.groupby('year').agg({"Sales" : "sum"}).reset_index()
df_y_sa['Sales'] = round(df_y_sa['Sales'],2)
df_y_sa['text'] = df_y_sa['Sales'].astype(str)

df_y_pr = df.groupby('year').agg({"Profit" : "sum"}).reset_index()
df_y_pr['Profit'] = round(df_y_pr['Profit'],2)
df_y_pr['text'] = df_y_pr['Profit'].astype(str)

Sales_growthrate = [0]
for i in range(3):
        Sales_growthrate.append(round((df_y_sa.iloc[i+1,1]-df_y_sa.iloc[i,1])/df_y_sa.iloc[i,1],4))
df_y_sa.insert(3,column="Sales_growthrate",value = Sales_growthrate)
df_y_sa['text2'] = (df_y_sa['Sales_growthrate']*100).astype(str)+"%"
df_y_sa['text2'][2] = "29.47%"

Profit_growthrate = [0]
for i in range(3):
        Profit_growthrate.append(round((df_y_pr.iloc[i+1,1]-df_y_pr.iloc[i,1])/df_y_pr.iloc[i,1],4))
df_y_pr.insert(3,column="Profit_growthrate",value = Profit_growthrate)
df_y_pr['text2'] = (df_y_pr['Profit_growthrate']*100).astype(str)+"%"


fig = make_subplots(rows=1, cols=3, vertical_spacing=0.08,
                    specs=[[{"type": "scatter"},None, {"type": "scatter"}]],
                    column_widths=[0.475,0.05, 0.475],
                    subplot_titles=("Sales vs Profit per year", 
                                    "growth rate Sales vs Profit per year"))

fig.add_trace(go.Scatter(x=df_y_sa['year'], y=df_y_sa['Sales'],mode = "lines+markers",line=dict(color="blue"),
                     text=df_y_sa['text'],name="Sales"),row=1, col=1)

fig.add_trace(go.Scatter(x=df_y_pr['year'], y=df_y_pr['Profit'],mode = "lines+markers",line=dict(color="red"),
                     text=df_y_pr['text'],name="Profit"),row=1, col=1)

fig.add_trace(go.Scatter(x=["2015","2016","2017"], y=df_y_sa['Sales_growthrate'][1:],mode = "lines+markers",line=dict(color="darkblue"),
                     text=df_y_sa['text2'][1:],name="Sales_growthrate"),row=1, col=3)

fig.add_trace(go.Scatter(x=["2015","2016","2017"], y=df_y_pr['Profit_growthrate'][1:],mode = "lines+markers",line=dict(color="darkred"),
                     text=df_y_pr['text2'][1:],name="Profit_growthrate"),row=1, col=3)

fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20), 
                  title_text="Sales vs Profit",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif"))

fig.show()


# ### 左圖為每年銷售與利潤總額比較，前兩年的營收雖然沒太多成長，但從 2016 年開始便有顯著成長，而利潤則是四年皆有穩定成長。
# 
# ### 右圖是銷售與利潤年成長率，在 2016 年的銷售額與獲利皆有成長，到了 2017 年雖然還保持著獲利，但兩項數據同時衰退，需深入探討原因。

# In[7]:


df_sc_sa = df.groupby("Sub-Category").agg({"Sales" : "mean"}).reset_index().sort_values(by="Sales", ascending=False)[:10]
df_ca_sa = df.groupby("Category").agg({"Sales" : "mean"}).reset_index().sort_values(by="Sales", ascending=False)

df_sc_sa["color"] = "#496595"
df_sc_sa["color"][2:] = "#c6ccd8"


fig = make_subplots(rows=1, cols=3, 
                    specs=[[{"type": "bar"},None,{"type": "pie"}]],
                    column_widths=[0.65,0.05,0.3], vertical_spacing=0, horizontal_spacing=0.02,
                    subplot_titles=("Top 10 Highest Sub-Category Sales","Highest Sales in Category"))

fig.add_trace(go.Pie(values=df_ca_sa["Sales"], labels=df_ca_sa["Category"], name="Category",
                     marker=dict(colors=["#334668","#496595","#6D83AA"]), hole=0.7,
                     hoverinfo="label+percent+value", textinfo="label"), 
                    row=1, col=3)
fig.add_trace(go.Bar(x=df_sc_sa["Sales"], y=df_sc_sa["Sub-Category"], marker=dict(color= df_sc_sa["color"]),
                     name="Sub-Category", orientation="h"), 
                     row=1, col=1)


fig.update_yaxes(showgrid=False, ticksuffix=" ", categoryorder="total ascending", row=1, col=1)
fig.update_xaxes(visible=False, row=1, col=1)
fig.update_layout(height=500, bargap=0.2,
                  margin=dict(b=0,r=20,l=20), xaxis=dict(tickmode="linear"),
                  title_text="Average Sales Analysis",
                  template="plotly_white",
                  title_font=dict(size=29, color="#8a8d93", family="Lato, sans-serif"),
                  font=dict(color="#8a8d93"),
                  showlegend=False)


fig.show()


# ### 左圖為以 "Sub-Category" 分類的平均客單價 AOV(Average Order Value)，並按照排名列出前十名，"Copiers" 和 "Machines" 兩子類別的平均客單價遠大於其他類別，也可發現前六名都是 "Technology" 和 "Furniture" 類別的產品。
# ### 右圖是以 "Category" 分類的AOV。在 "OfficeSupplies" 分類的平均客單價明顯低於其他兩個分類，若想提升此類別之AOV，可嘗試提供大包裝設計或加購活動。

# In[8]:


df_2014 = df[df['year']==2014][['month','Sales']]
df_2014 = df_2014.groupby('month').agg({"Sales" : "mean"}).reset_index().rename(columns={'Sales':'s14'})
df_2015 = df[df['year']==2015][['month','Sales']]
df_2015 = df_2015.groupby('month').agg({"Sales" : "mean"}).reset_index().rename(columns={'Sales':'s15'})
df_2016 = df[df['year']==2016][['month','Sales']]
df_2016 = df_2016.groupby('month').agg({"Sales" : "mean"}).reset_index().rename(columns={'Sales':'s16'})
df_2017 = df[df['year']==2017][['month','Sales']]
df_2017 = df_2017.groupby('month').agg({"Sales" : "mean"}).reset_index().rename(columns={'Sales':'s17'})
df_year = df_2014.merge(df_2015,on='month').merge(df_2016,on='month').merge(df_2017,on='month')


top_labels = ['2014', '2015', '2016', '2017']


colors = ['rgba(38, 24, 74, 0.8)', 'rgba(71, 58, 131, 0.8)',
          'rgba(122, 120, 168, 0.8)', 'rgba(164, 163, 204, 0.85)']

df_year = df_year[['s14','s15','s16','s17']].replace(np.nan,0)
x_data = df_year.values


df_2014['month'] =['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
y_data = df_2014['month'].tolist()

fig = go.Figure()
for i in range(0, len(x_data[0])):
    for xd, yd in zip(x_data, y_data):
        fig.add_trace(go.Bar(
            x=[xd[i]], y=[yd],
            orientation='h',
            marker=dict(
                color=colors[i],
                line=dict(color='rgb(248, 248, 249)', width=1)
            )
        ))

fig.update_layout(title='Avg Sales for each Year',
    xaxis=dict(showgrid=False, 
               zeroline=False, domain=[0.15, 1]),
    yaxis=dict(showgrid=False, showline=False,
               showticklabels=False, zeroline=False),
    barmode='stack', 
    template="plotly_white",
    margin=dict(l=0, r=50, t=100, b=10),
    showlegend=False, 
)

annotations = []
for yd, xd in zip(y_data, x_data):
    # labeling the y-axis
    annotations.append(dict(xref='paper', yref='y',
                            x=0.14, y=yd,
                            xanchor='right',
                            text=str(yd),
                            font=dict(family='Arial', size=14,
                                      color='rgb(67, 67, 67)'),
                            showarrow=False, align='right'))
    # labeling the first Likert scale (on the top)
    if yd == y_data[-1]:
        annotations.append(dict(xref='x', yref='paper',
                                x=xd[0] / 2, y=1.1,
                                text=top_labels[0],
                                font=dict(family='Arial', size=14,
                                          color='rgb(67, 67, 67)'),
                          showarrow=False))
    space = xd[0]
    for i in range(1, len(xd)):
            # labeling the Likert scale
            if yd == y_data[-1]:
                annotations.append(dict(xref='x', yref='paper',
                                        x=space + (xd[i]/2), y=1.1,
                                        text=top_labels[i],
                                        font=dict(family='Arial', size=14,
                                                  color='rgb(67, 67, 67)'),
                                        showarrow=False))
            space += xd[i]
fig.update_layout(
    annotations=annotations)
fig.show()


# ### 上圖為各月份的平均每筆訂單銷售額，並以深淺色區分不同年份的資料。可初步看出二月份銷量最低、三月份銷量最高，此現象在2014年尤其明顯。

# In[9]:


df_m_sa = df.groupby('month').agg({"Sales" : "mean"}).reset_index()
df_m_sa['Sales'] = round(df_m_sa['Sales'],2)
df_m_sa['month_text'] = df_m_sa['month'].apply(lambda x: calendar.month_abbr[x])
df_m_sa['text'] = df_m_sa['month_text'] + ' - ' + df_m_sa['Sales'].astype(str) 

df_dw_sa = df.groupby('day_of_week').agg({"Sales" : "mean"}).reset_index() 
df_dw_sa = df_dw_sa.reindex([3,2,0,4,6,5,1]).reset_index()
df_dw_sa['Sales'] = round(df_dw_sa['Sales'], 2)
df_dw_sa['text'] = df_dw_sa['Sales'].astype(str) 

df_m_sa['color'] = '#c6ccd8'
df_m_sa['color'][2] = '#496595'
df_dw_sa['color'] = '#c6ccd8'
df_dw_sa['color'][5] = '#496595'

fig = make_subplots(rows=1, cols=3, vertical_spacing=0.08,
                    specs=[[{"type": "bar"},None, {"type": "bar"}]],
                    column_widths=[0.475,0.05, 0.475],
                    subplot_titles=("Month wise Avg Sales Analysis", 
                                    "Avg Sales vs Day of Week"))

fig.add_trace(go.Bar(x=df_m_sa['Sales'], y=df_m_sa['month'], marker=dict(color= df_m_sa['color']),
                     text=df_m_sa['text'],textposition='auto',
                     name='Month', orientation='h'), 
                     row=1, col=1)

fig.add_trace(go.Bar(x=df_dw_sa['Sales'], y=df_dw_sa['day_of_week'], marker=dict(color= df_dw_sa['color']),
                     text=df_dw_sa['text'],textposition='auto',
                     name='day_of_week', orientation='h'), 
                     row=1, col=3)


fig.update_yaxes(visible=False, row=1, col=1)
fig.update_xaxes(visible=False, row=1, col=1)
fig.update_xaxes(visible=False, row=1, col=3)
fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20), 
                  title_text="Average Sales Analysis",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif"),
                  showlegend=False)

fig.show()


# ### 左圖為四年之間各月份的平均銷售額，可看出三月的平均每筆訂單銷售額最高，二月則為最低，可對照行銷活動舉辦月份，以確認活動是否有達到成效。
# ### 右圖則是以一週中的星期幾作為分類依據，星期二為顧客最願意花錢購買商品的日子，而週末的AOV則表現平平。

# In[10]:


df_new_or_reg = df.groupby(["Order ID","year",'new_customer']).agg({"Sales" : "count"}).reset_index()
df_new_or_reg = df_new_or_reg.groupby(["year",'new_customer']).agg({"Sales" : "count"}).reset_index()

df_new_or_reg["year"]= df_new_or_reg["year"].astype("str")
df_new_or_reg["new_customer"]= df_new_or_reg["new_customer"].astype("str")

fig = px.bar(df_new_or_reg, x="year", y="Sales",color="new_customer", color_discrete_sequence=['#496595', '#c6ccd8'],
            labels={"Sales": "counts"})


fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20),title_text="new_customer vs regular_customer",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif"))

fig.show()


# ### 上圖為各年新客和常客的訂單數比較，由客戶組成可看出每年訂單數皆有成長、且有許多常客會回頭購買，但新客訂單數衰退地十分嚴重，到2017年新客只剩下28筆訂單，可能需找尋吸引新顧客的方法。

# In[11]:


df_dis_or_not = df.groupby(["Order ID","year",'discount_ornot']).agg({"Sales" : "count"}).reset_index()
df_dis_or_not = df_dis_or_not.groupby(["year",'discount_ornot']).agg({"Sales" : "count"}).reset_index()

df_dis_or_not["year"]= df_dis_or_not["year"].astype("str")
df_dis_or_not["discount_ornot"]= df_dis_or_not["discount_ornot"].astype("str")


fig = px.bar(df_dis_or_not, x="year", y="Sales",color="discount_ornot", color_discrete_sequence=['#496595', '#c6ccd8'],
            labels={"Sales": "order numbers"})

fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20),title_text="Discounted vs Original price",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif"))


fig.show()


# ### 上圖為各年訂單為有折扣或是原價購買之比較，兩類訂單數量差距不大但可看出有折扣的數量略多一些，可接著由銷售額與獲利探討詳細情形。

# In[12]:


df_dis_sa = df.groupby(["Order Date"]).agg({"Sales" : "sum"}).reset_index().rename(columns={'Sales':'num'})
df_dis_pr = df.groupby(["Order Date"]).agg({"Profit" : "sum"}).reset_index().rename(columns={'Profit':'num'})

fig = go.Figure()
fig.add_trace(go.Scatter(x=df_dis_sa["Order Date"], y=df_dis_sa['num'],name="Sales",line=dict(color="lightskyblue")))
fig.add_trace(go.Scatter(x=df_dis_pr["Order Date"], y=df_dis_pr['num'],name="Profit",line=dict(color="red")))


fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20),title_text="Sales vs Profit",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif")
)

fig.show()


# ### 上圖為四年期間每天的營業額與利潤比較，圖中看到2014/3/18的營業額非常高但沒有反映在利潤上，而2015年之前營業額出現的幾個高峰也都沒帶來相對的利潤，可認為是為吸引新顧客舉辦的活動等等，在2016年後營業額便有與利潤同步的趨勢，但在2016/11/25卻出現虧損的峰值，需特別留意造成此狀況的原因。

# In[13]:


df_dis = df[df["discount_ornot"]==1]
df_dis_sa = df_dis.groupby(["Order Date"]).agg({"Sales" : "sum"}).reset_index().rename(columns={'Sales':'num'})
df_dis_pr = df_dis.groupby(["Order Date"]).agg({"Profit" : "sum"}).reset_index().rename(columns={'Profit':'num'})


fig = go.Figure()
fig.add_trace(go.Scatter(x=df_dis_sa["Order Date"], y=df_dis_sa['num'],name="Sales",line=dict(color="lightskyblue")))
fig.add_trace(go.Scatter(x=df_dis_pr["Order Date"], y=df_dis_pr['num'],name="Profit",line=dict(color="red")))


fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20),title_text="Discounted order : Sales vs Profit",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif")
)

fig.show()


# ### 此為有折扣的訂單，可確認2014/3/18的銷量主要為折扣的訂單組成，可推測是舉辦折扣活動吸引新客，也發現2016/11/25造成的虧損為折扣訂單導致，需再詳細確認詳細狀況。

# In[14]:


df_nodis = df[df["discount_ornot"]==0]
df_dis_sa = df_nodis.groupby(["Order Date"]).agg({"Sales" : "sum"}).reset_index().rename(columns={'Sales':'num'})
df_dis_pr = df_nodis.groupby(["Order Date"]).agg({"Profit" : "sum"}).reset_index().rename(columns={'Profit':'num'})


fig = go.Figure()
fig.add_trace(go.Scatter(x=df_dis_sa["Order Date"], y=df_dis_sa['num'],name="Sales",line=dict(color="lightskyblue")))
fig.add_trace(go.Scatter(x=df_dis_pr["Order Date"], y=df_dis_pr['num'],name="Profit",line=dict(color="red")))


fig.update_layout(height=500, bargap=0.15,
                  margin=dict(b=0,r=20,l=20),title_text="Original price order : Sales vs Profit",
                  template="plotly_white",
                  title_font=dict(size=29, color='#8a8d93', family="Lato, sans-serif"),
                  font=dict(color='#8a8d93'),
                  hoverlabel=dict(bgcolor="#f2f2f2", font_size=13, font_family="Lato, sans-serif")
)

fig.show()


# ### 最後一張圖可確認2016/10/2與2017/3/23的營業額與利潤皆高，可確認這兩天主要訂單為原價訂單，對照上方折扣訂單在這兩天表現平平，可推測為熱銷新品開賣的日子。

# In[15]:


# df.to_csv("/Users/yoyo/Downloads/superstore/superstore_update.csv")

