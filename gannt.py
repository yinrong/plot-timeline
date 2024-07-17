import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

# 从CSV文件中读取项目数据
df = pd.read_csv('projects.csv')

# 将日期转换为datetime对象
df['开始时间'] = pd.to_datetime(df['开始时间'], format='%Y%m%d')
df['结束时间'] = pd.to_datetime(df['结束时间'], format='%Y%m%d')

# 填充模块的空白值
df['模块'] = df['模块'].ffill()

# 生成自定义的x轴标签
x_ticks = pd.date_range(start=df['开始时间'].min(), end=df['结束时间'].max(), freq='MS')
x_labels = []
previous_year = None

for i, tick in enumerate(x_ticks):
    if tick.month == 1 or tick.year != previous_year or i == 0:
        x_labels.append(tick.strftime('%Y-%m'))
        previous_year = tick.year
    else:
        x_labels.append(tick.strftime('%m'))

# 填充空白的开始时间
for i in range(1, len(df)):
    if pd.isna(df.at[i, '开始时间']):
        previous_end_date = df.at[i-1, '结束时间']
        new_start_date = previous_end_date + timedelta(days=1)
        df.at[i, '开始时间'] = new_start_date

# 获取所有负责人的唯一名称
unique_responsibles = df['负责人'].dropna().unique()

# 生成颜色映射
colors = px.colors.qualitative.Plotly
color_mapping = {responsible: colors[i % len(colors)] for i, responsible in enumerate(unique_responsibles)}

# 绘制甘特图
fig = px.timeline(
    df,
    x_start="开始时间",
    x_end="结束时间",
    y="关键里程碑",
    color="负责人",
    labels={
        "开始时间": "开始日期",
        "结束时间": "结束日期",
        "关键里程碑": "里程碑",
        "负责人": "负责人",
        "模块": "模块"
    },
    color_discrete_map=color_mapping,
    category_orders={"关键里程碑": df["关键里程碑"].tolist()},  # 强制按CSV顺序排序
)

# 更新布局以适应更好的显示效果
fig.update_layout(
    xaxis_title="日期",
    yaxis_title="里程碑",
    font=dict(size=15),
    title_font_size=20,
    xaxis_tickformat="%m\n%Y",
    xaxis=dict(
        tickmode='array',
        tickvals=x_ticks,
        ticktext=x_labels,
    ),
    annotations=[]  # 清空默认注释以添加自定义注释
)

# 添加每个模块的注释和虚线贯穿整个图表
for module in df['模块'].unique():
    module_tasks = df[df['模块'] == module]
    module_end = module_tasks['结束时间'].max()
    module_text = module.replace('\n', '<br>')  # 确保显示换行符
    print(module_text, module_end)
    fig.add_annotation(
        x=module_end,
        y=-0.01,  # 向 y 轴负方向移动 0.01
        yref="paper",
        showarrow=True,
        arrowhead=2,
        ax=module_end,
        axref='x',
        ay=60,  # 确保箭头指向时间轴
        ayref='y',
        arrowcolor="rgba(255, 160, 122, 0.5)",
        arrowsize=2,
        arrowwidth=2,
        text=module_text,
        font=dict(size=14),
        bgcolor="LightSalmon"
    )
    fig.add_vline(x=module_end, line_width=1, line_dash="dash", line_color="LightSalmon")

# 获取每个模块的起始和结束位置
module_end_positions = df.groupby('模块').apply(lambda x: x.index[-1]).tolist()

for i in range(16):
    fig.add_hline(y=i-0.5, line_width=0.5, line_dash="dash", line_color="LightSalmon")
for end_pos in module_end_positions:
    y = len(df) - 1 - end_pos - 0.5
    fig.add_hline(y=y, line_width=2, line_dash="dash", line_color="LightSalmon")


# 将图表保存为 HTML 文件
fig.write_html("gantt_chart.html")
