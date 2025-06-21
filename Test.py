import pandas as pd

# ========= 1. 参数 ========= #
SRC_FILE  = r'C:\Users\Ghost\Desktop\充电桩AI分析\23-24年小区公共桩用户数据.xlsx'      # 原始 Excel 路径
DST_FILE  = r'C:\Users\Ghost\Desktop\充电桩AI分析\正式数据.xlsx'  # 输出文件（可改 .xlsx / .csv）

# ========= 2. 读取原始宽表 ========= #
# 跳过第 2 行（Excel 序列号），把第 1 & 3 行当作两级列索引
df = pd.read_excel(
        SRC_FILE,
        header=[0, 1],
        skiprows=[1],        # ← 关键
        dtype={'户号': str},  # 保护户号前导 0
        engine='openpyxl'
     )

# ========= 3. 清洗表头 / 删掉「序号」列 ========= #
seq_cols = df.columns[df.columns.get_level_values(0) == '序号']
df = df.drop(columns=seq_cols)               # 去序号
# 把 (日期, 指标) → f"{YYYY-MM-DD}_{指标}"，户号直接叫 '户号'
new_cols = []
for lvl0, lvl1 in df.columns:
    if lvl0 == '户号':
        new_cols.append('户号')
    else:
        new_cols.append(f"{lvl0:%Y-%m-%d}_{lvl1}")  # 2023-01-01_充电量（度）
df.columns = new_cols

# ========= 4. 宽表 → 长表（melt） ========= #
long = (
    df.melt(id_vars='户号', var_name='日期_指标', value_name='值')
      .assign(  # 拆分列名
          日期   = lambda d: d['日期_指标'].str.split('_').str[0],
          指标   = lambda d: d['日期_指标'].str.split('_').str[1]
      )
      .drop(columns='日期_指标')
)

# ========= 5. 指标列再 Pivot 成两列 ========= #
tidy = (long
        .pivot_table(index=['户号', '日期'],
                     columns='指标',
                     values='值',
                     aggfunc='first')
        .reset_index()
)

# ========= 6. 周几（中文）& 列名统一 ========= #
weekday_map = {0:'星期一',1:'星期二',2:'星期三',
               3:'星期四',4:'星期五',5:'星期六',6:'星期日'}
tidy['周几'] = pd.to_datetime(tidy['日期']).dt.dayofweek.map(weekday_map)

tidy = tidy.rename(columns={
    '充电量（度）':'充电量(度)',
    '充电时长（分）':'充电时长(分)'
})

# 确保数值列是 float / int
for col in ['充电量(度)', '充电时长(分)']:
    tidy[col] = pd.to_numeric(tidy[col], errors='coerce')

# ========= 6.1 过滤掉“双 0”记录（新增） ========= #
# 充电量和充电时长同时为 0（或 NaN→0）就剔除
tidy = tidy.fillna(0)  # 若不想把 NaN 当 0，可删掉这一行
tidy = tidy.loc[(tidy['充电量(度)'] != 0) | (tidy['充电时长(分)'] != 0)]

# ========= 6.5 重新排列顺序：把“周几”挪到“日期”后面 ========= #
cols_order = ['户号', '日期', '周几', '充电量(度)', '充电时长(分)']
tidy = tidy[cols_order]

# ========= 7. 保存 ========= #
# tidy.to_parquet(DST_FILE)          # 体积最小 & 读写最快
tidy.to_excel(DST_FILE, index=False)  # 如需 Excel
# tidy.to_csv('charge_tidy.csv',  index=False)    # 如需 CSV

print(f'整理完成，共 {len(tidy):,} 行，已保存到 {DST_FILE}')
