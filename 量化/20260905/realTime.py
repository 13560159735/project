# 腾讯行情接口-实时数据

# 接口：https://qt.gtimg.cn/q={symbol}

# 参数
#     symbol:替换成带市场前缀的股票代码，如 sh600519 (贵州茅台)

# resp.text之后，返回字符串，以sh600519(贵州茅台)为例    v_sh600519="1~贵州茅台~600519~1291.79~1307.88~1299.80~15318~6402~8909~1291.35~2~1291.33~1~1291.30~5~1291.28~1~1291.26~2~1291.86~2~1291.87~4~1292.00~1~1292.20~1~1292.46~1~~20260820112228~-16.09~-1.23~1306.88~1291.00~1291.79/15318/1986540868~15318~198654~0.12~19.83~~1306.88~1291.00~1.21~16148.43~16148.43~6.43~1438.67~1177.09~0.75~2~1296.85~18.14~19.62~~~0.15~198654.0868~0.0000~0~   A~GP-A~-4.25~-4.69~4.03~32.41~27.30~1539.98~1151.01~-1.28~-0.02~1.32~1250081601~1250081601~10.00~-7.73~1250081601~~~-7.60~-0.00~~CNY~0~___D__F__N~1291.20~13~";

# 结构说明：
# 整体是：v_股票代码="~分隔的字符串"；
# 代码用正则取双引号内的部分，再按 ~ split 成列表 text_split_value (共88个元素)
# 常用字段下标对照（text_split_value[下标] -> 含义）：
# text_split_value[0] 未知/类别（1）
# text_split_value[1] 名称（贵州茅台）
# text_split_value[2] 代码（600519）
# text_split_value[3] 当前价（1291.79）
# text_split_value[4] 昨收（1307.88）
# text_split_value[5] 今开（1299.80）
# text_split_value[6] 成交量（手）
# text_split_value[7] 外盘
# text_split_value[8] 内盘
# text_split_value[9] ~ text_split_value[28] 买一~买五/卖一~卖五（价格~数量 成对出现）
# text_split_value[30] 行情时间（yyyymmddHHMMSS,如20260820112228)
# text_split_value[31] 涨跌额（-16.09）
# text_split_value[32] 涨跌幅%（-1.23）
# text_split_value[33] 最高（1306.88）
# text_split_value[34] 最低（1291.00）
# text_split_value[36] 成交量（手，15318）
# text_split_value[37] 成交额（万元，198654）

import datetime
import re
import time
import requests

#用户输入股票代码，规范为接口所需的市场前缀股票代码，如600519->sh600519
def normalize_code(raw:str):
    code=raw.strip().lower()
    if re.fullmatch(r"(sh|sz|bj)\d{6}",code):
        return code


