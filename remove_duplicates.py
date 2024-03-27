import pandas as pd

# 读取xlsx文件
df = pd.read_excel('your_file.xlsx')

# 去除H列中各个单元格中的重复词
df['Products'] = df['Products'].apply(lambda x: ', '.join(set(str(x).split(', '))))

# 保存结果到新的xlsx文件
df.to_excel('updated_file.xlsx', index=False)

print("重复词去除完成。")
