import pandas as pd  # 导入pandas库用于数据处理
import numpy as np  # 导入numpy库用于数值计算
import plotly.graph_objects as go  # 导入plotly.graph_objects库用于绘图
from datetime import datetime  # 导入datetime库用于日期时间处理
################################################################################

# K线数据EXCEL文件路径
####################### 请修改为你的文件路径 #######################
file_path = r'C:\Users\yhqhzz\Desktop\股指数据\3月_交易信号标注\options_data\2503\3M-P-4000-3950.xlsx'
# 交易信号EXCEL文件路径
####################### 请修改为你的文件路径 #######################
signal_path = r'C:\Users\yhqhzz\Desktop\股指数据\3月_交易信号标注\trading_signal_data\账户一\3M-P-4000-3950-20250225-20250314.xlsx'
# K线图存放路径
####################### 请修改为你的文件路径 #######################
chart_path = r'C:\Users\yhqhzz\Desktop\股指数据\3月_交易信号标注\chart\账户一'
# K线图合约名称
####################### 请修改为实际合约 #######################
contract_name = '3M-P-4000-3950_1'
# 定义目标时间段
####################### 请修改为实际时间段 #######################
start_date = pd.Timestamp('2025-02-20')
end_date = pd.Timestamp('2025-03-20')
# 截取部分K线图，加快出图速度
################################################################################




# 读取数据并处理时间
df = pd.read_excel(file_path)
df['timestamp'] = pd.to_datetime(df['交易时间'], errors='coerce')
df = df.dropna(subset=['timestamp']).set_index('timestamp')
# 检查数据时间段
print("数据时间范围:", df.index.min(), "至", df.index.max())
# 筛选数据
filtered_df = df.loc[start_date:end_date]

if filtered_df.empty:
    print("错误：指定时间段内无数据！")
else:
    # 数据聚合处理
    ohlc = filtered_df['收盘价'].resample('5min').ohlc()
    ohlc.columns = ['Open', 'High', 'Low', 'Close']
    kline_data = ohlc.dropna().reset_index()
    kline_data['x_index'] = range(len(kline_data))
    kline_data['timestamp'] = kline_data['timestamp'].dt.floor('5min')

    # 创建图表
    fig = go.Figure(data=[go.Candlestick(  
        x=kline_data['x_index'],  # 使用 x_index 作为 x 轴
        open=kline_data['Open'],  # 设置开盘价
        close=kline_data['Close'],  # 设置收盘价
        high=kline_data['High'],    # 设置最高价
        low=kline_data['Low'],      # 设置最低价
        increasing={  # 设置上涨蜡烛的颜色和填充色
            'line': {'color': 'rgb(255,113,113)'},  # 边框颜色
            'fillcolor': 'rgba(255,113,113,0.4)'  # 填充颜色（带透明度）
        },
        decreasing={  # 设置下跌蜡烛的颜色和填充色
            'line': {'color': 'green'},  # 边框颜色
            'fillcolor': 'rgba(0,128,0,0.4)'  # 填充颜色（带透明度）
        },
        hoverinfo='text'
    )])

    # 统计买开仓、卖开仓、卖平仓、买平仓的数据
    trade_signals = pd.read_excel(signal_path)
    signal_records = {}  # 存储每个时间点的交易信号
    buy_open_signals = []
    sell_open_signals = []
    buy_close_signals = []
    sell_close_signals = []

    for _, signal in trade_signals.iterrows():
        signal_time = pd.to_datetime(signal['交易时间'])
        if start_date <= signal_time <= end_date:
            kline_time = signal_time.floor('5min')
            kline_row = kline_data[kline_data['timestamp'] == kline_time]
        
            if not kline_row.empty:
                x_pos = kline_row['x_index'].values[0]
                candle_low = kline_row['Low'].values[0]
                candle_high = kline_row['High'].values[0]
                close_price = kline_row['Close'].values[0]
            
                if signal['买卖方向'] == '买开仓':
                    fig.add_annotation(
                        x=x_pos, 
                        y=candle_low - 0.2,
                        text='▲',
                        showarrow=False,
                        font=dict(color='red', size=14),
                        yshift=-15,
                        yref='y'
                    )
                    buy_open_signals.append((x_pos, close_price))
                elif signal['买卖方向'] == '卖开仓':
                    fig.add_annotation(
                        x=x_pos,
                        y=candle_high + 0.2,
                        text='▼',
                        showarrow=False,
                        font=dict(color='green', size=14),
                        yshift=15,
                        yref='y'
                    )
                    sell_open_signals.append((x_pos, close_price))
                elif signal['买卖方向'] == '买平仓':
                    fig.add_annotation(
                        x=x_pos, 
                        y=candle_low - 0.2,
                        text='▲',
                        showarrow=False,
                        font=dict(color='red', size=14),
                        yshift=-15,
                        yref='y'
                    )
                    buy_close_signals.append((x_pos, close_price))
                elif signal['买卖方向'] == '卖平仓':
                    fig.add_annotation(
                    x=x_pos,
                    y=candle_high + 0.2,
                    text='▼',
                    showarrow=False,
                    font=dict(color='green', size=14),
                    yshift=15,
                    yref='y'
                    )
                    sell_close_signals.append((x_pos, close_price))
            
                # 记录交易信号
                if kline_time not in signal_records:
                    signal_records[kline_time] = []
                signal_records[kline_time].append({'direction': signal['买卖方向'], 'price': signal['价格']})


    # 设置 hovertemplate 以显示具体时间
    hover_templates = []
    for idx, row in kline_data.iterrows():
        hover_text = f"时间: {row['timestamp']}<br>开盘: {row['Open']:.2f}<br>最高: {row['High']:.2f}<br>最低: {row['Low']:.2f}<br>收盘: {row['Close']:.2f}"
        
        # 添加交易信号信息
        if row['timestamp'] in signal_records:
            for signal in signal_records[row['timestamp']]:
                hover_text += f"<br>交易信号: {signal['direction']} @ {signal['price']:.2f}"
        
        hover_templates.append(hover_text)
    
    fig.data[0].hovertext = hover_templates  # 设置Candlestick图表的hover模板

    # 连接买开仓和卖平仓信号
    connection_count_b = 0  # 新增：连接线计数器
    sell_signal_indices = list(range(len(sell_close_signals)))  # 修改：使用正确的sell_close_signals
    
    for buy_signal in buy_open_signals:
        closest_sell_signal = None
        min_distance = float('inf')
        for i in sell_signal_indices:
            if sell_close_signals[i][0] > buy_signal[0]:  # 修改：使用正确的sell_close_signals
                distance = sell_close_signals[i][0] - buy_signal[0]  # 修改：使用正确的sell_close_signals
                if distance < min_distance:
                    min_distance = distance
                    closest_sell_signal = sell_close_signals[i]  # 修改：使用正确的sell_close_signals
                    closest_sell_index = i
        
        if closest_sell_signal:
            fig.add_shape(
                type='line',
                x0=buy_signal[0], y0=buy_signal[1],
                x1=closest_sell_signal[0], y1=closest_sell_signal[1],
                line=dict(color='red', dash='dash', width=1)
            )
            connection_count_b += 1  # 新增：计数成功连接
            sell_signal_indices.remove(closest_sell_index)

    # 连接卖开仓和买平仓信号
    connection_count_s = 0  # 新增：连接线计数器
    sell_open_indices = list(range(len(sell_open_signals)))  # 修改：变量名修正

    for buy_close_signal in buy_close_signals:
        closest_sell_open_signal = None
        min_distance = float('inf')
        for i in sell_open_indices:
            if sell_open_signals[i][0] < buy_close_signal[0]:
                distance = buy_close_signal[0] - sell_open_signals[i][0]
                if distance < min_distance:
                    min_distance = distance
                    closest_sell_open_signal = sell_open_signals[i]
                    closest_sell_open_index = i
    
        if closest_sell_open_signal:
            fig.add_shape(
                type='line',
                x0=closest_sell_open_signal[0], y0=closest_sell_open_signal[1],
                x1=buy_close_signal[0], y1=buy_close_signal[1],
                line=dict(color='green', dash='dash', width=1)
            )
            connection_count_s += 1  # 新增：计数成功连接
            sell_open_indices.remove(closest_sell_open_index)

    # 新增统计代码
    stats_df = pd.DataFrame({
        '类型': ['买开仓', '卖平仓', '卖开仓', '买平仓', '多头连接线', '空头连接线'],
        '数量': [len(buy_open_signals), len(sell_close_signals), len(sell_open_signals), len(buy_close_signals), connection_count_b, connection_count_s]
    })
    print("\n交易信号统计：")
    print(stats_df.to_string(index=False))
        
    # --- Y轴标识线配置 ---
    min_price = kline_data['Low'].min()
    max_price = kline_data['High'].max()
    price_range = max_price - min_price
    
    # 动态计算间隔
    if price_range <= 5:  # 小波动行情
        base_interval = 0.5
        major_interval = 1
    elif 5 < price_range <= 20:  # 中等波动
        base_interval = 1
        major_interval = 2
    else:  # 大波动行情
        base_interval = 2
        major_interval = 5
    
    # 基础网格线（动态间隔）
    base_levels = np.arange(np.floor(min_price/base_interval)*base_interval, 
                           np.ceil(max_price/base_interval)*base_interval + base_interval, 
                           base_interval)
    for level in base_levels:
        fig.add_hline(
            y=level,
            line=dict(color='rgba(200,200,200,0.2)', width=0.5),
            layer='below'
        )

    # 重点网格线（动态间隔）
    major_levels = np.arange(np.floor(min_price/major_interval)*major_interval, 
                            np.ceil(max_price/major_interval)*major_interval + major_interval, 
                            major_interval)
    for level in major_levels:
        fig.add_hline(
            y=level,
            line=dict(color='rgba(150,150,150,0.5)', width=1, dash='dot'),
            annotation_text=f"{level:.2f}",
            annotation_position="right"
        )

    # --- X轴配置 ---
    daily_markers = ['10:00', '11:30', '14:00', '15:00']
    unique_dates = kline_data['timestamp'].dt.date.unique()
    tickvals, ticktext = [], []

    for date in unique_dates:
        for time_str in daily_markers:
            full_time = pd.to_datetime(f"{date} {time_str}")
            if full_time in kline_data['timestamp'].values:
                x_pos = kline_data.loc[kline_data['timestamp'] == full_time, 'x_index'].values[0]
                tickvals.append(x_pos)
                ticktext.append(full_time.strftime('%m-%d %H:%M'))

    fig.update_xaxes(
        tickvals=tickvals,
        ticktext=ticktext,
        tickangle=45,
        title_text='时间',
        tickfont=dict(size=10),
        showgrid=False
    )

    # 每日起始垂直线
    daily_starts = kline_data.groupby(kline_data['timestamp'].dt.date)['x_index'].min()
    for x_pos in daily_starts:
        fig.add_vline(
            x=x_pos,
            line=dict(color='rgba(150, 150, 150, 0.5)', width=1, dash='dot'),
            layer='below'
        )

    # --- 最终布局配置 ---
    fig.update_layout(
        title=f'{contract_name} ({start_date.date()} 至 {end_date.date()})',
        yaxis_title='价格',
        xaxis=dict(
            showline=True,
            linecolor='black',
            rangeslider=dict(
                visible=True,
                thickness=0.05,
                bgcolor='rgba(150,150,150,0.2)'
            )
        ),
        plot_bgcolor='white',
        yaxis=dict(
            autorange=True,
            fixedrange=False
        ),
        # 新增交互配置
        dragmode='pan',  # 拖拽平移
        hovermode='x unified',  # 显示垂直参考线
        # 启用以下高级交互功能
        meta=dict(
            plotly={
                'scrollZoom': True,  # 启用滚轮缩放
                'displayModeBar': True,  # 显示模式栏
                'modeBarButtonsToAdd': ['hoverclosest', 'hovercompare'],
                'displaylogo': False
            }
        )
    )

    # 导出文件
    fig.write_html(f"{chart_path}\\{contract_name}.html", include_plotlyjs='cdn')
    print("HTML文件已生成！")
    print(f"文件生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")  # 新增时间打印