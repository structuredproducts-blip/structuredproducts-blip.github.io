# -*- coding: utf-8 -*-
"""FCN 榜单中文化词表:公司名 / 行业(yfinance子板块)英文→中文。
用于 rebuild_lists.py 每周重建时把英文名称/行业转中文。
- 已是中文的原样保留;
- 无通用中文名的纯品牌(Cloudflare/Coinbase/Shopify/DoorDash/Circle/Nebius/Workday/MiniMax 等)故意不入表,保留英文(中文财经惯例)。
- 新出现的英文名如需中文,往对应字典里加一行即可。
"""

def _is_cn(s):
    return any('一' <= c <= '鿿' for c in str(s or ""))

# 公司英文常用名 → 中文(仅收录有通用中文名的)
NAME_MAP = {
    "AMD": "超微半导体",
    "Arista": "阿里斯塔", "Arista Networks": "阿里斯塔",
    "Nu Holdings": "Nu控股",
    "MercadoLibre": "美客多",
    "Meta": "Meta平台", "Meta Platforms": "Meta平台",
    "Monolithic Power Systems": "芯源系统",
    "Marvell": "迈威尔科技", "Marvell Technology": "迈威尔科技",
    "Palantir": "帕兰提尔", "Palantir Technologies": "帕兰提尔",
    "Robinhood": "罗宾汉",
    "Roblox": "罗布乐思",
    "SK Hynix": "SK海力士",
    "Spotify": "声田", "Spotify Technology": "声田",
    "Cheniere Energy": "切尼尔能源",
    "Astera Labs": "阿斯特拉实验室",
    "Carvana": "卡瓦纳",
    "Constellation Energy": "星座能源",
    "ASMPT": "ASM太平洋",
    # 常见大盘(万一 universe 里是英文)
    "NVIDIA": "英伟达", "Nvidia": "英伟达",
    "Advanced Micro Devices": "超微半导体",
    "Broadcom": "博通", "Micron": "美光", "Micron Technology": "美光",
    "Seagate": "希捷科技", "Seagate Technology": "希捷科技",
    "Western Digital": "西部数据", "SanDisk": "闪迪",
    "Dell": "戴尔科技", "Dell Technologies": "戴尔科技",
    "HP Enterprise": "慧与科技", "Hewlett Packard Enterprise": "慧与科技",
    "Teradyne": "泰瑞达", "Applied Materials": "应用材料",
    "Lam Research": "泛林集团", "KLA": "科磊",
    "Texas Instruments": "德州仪器", "Analog Devices": "亚德诺",
    "NXP Semiconductors": "恩智浦", "Amphenol": "安费诺",
    "Taiwan Semiconductor": "台积电", "TSMC": "台积电", "ASML": "阿斯麦",
    "Salesforce": "赛富时", "ServiceNow": "ServiceNow", "Workday": "Workday",
    "Amazon": "亚马逊", "Apple": "苹果", "Microsoft": "微软",
    "Alphabet": "谷歌", "Meta Platforms Inc": "Meta平台",
    "Netflix": "奈飞", "Uber": "优步", "DoorDash": "DoorDash",
    "Newmont": "纽曼矿业", "Freeport-McMoRan": "麦克莫兰铜金",
    "Valero": "瓦莱罗能源", "Valero Energy": "瓦莱罗能源",
    "Chevron": "雪佛龙", "ExxonMobil": "埃克森美孚", "ConocoPhillips": "康菲石油",
    "Schlumberger": "斯伦贝谢", "Gilead": "吉利德科学", "Gilead Sciences": "吉利德科学",
    "Merck": "默沙东", "Bristol Myers Squibb": "施贵宝", "Amgen": "安进",
    "Target": "塔吉特", "Walmart": "沃尔玛", "Costco": "好市多",
    "GE Vernova": "GE Vernova", "Constellation": "星座能源",
    "Coinbase": "Coinbase", "Crowdstrike": "CrowdStrike",
    "DoorDash Inc": "DoorDash",
    "Oracle": "甲骨文", "Applied Optoelectronics": "应用光电",
    "Lumentum": "Lumentum", "Ciena": "Ciena", "Shopify": "Shopify",
    "Cloudflare": "Cloudflare", "Circle": "Circle", "Nebius": "Nebius",
    "Booking Holdings": "Booking", "DoorDash, Inc.": "DoorDash",
    "Coinbase Global": "Coinbase", "Nu Holdings Ltd": "Nu控股",
}

# yfinance 子板块 → 中文
IND_MAP = {
    "AI Software": "AI软件",
    "Aerospace & Defense": "航空航天与国防",
    "Apparel Retail": "服装零售",
    "Auto & Truck Dealerships": "汽车与卡车经销",
    "Auto Manufacturers": "汽车制造",
    "Auto Parts": "汽车零部件",
    "Banks - Diversified": "综合性银行",
    "Banks - Regional": "区域性银行",
    "Beverages - Non-Alcoholic": "非酒精饮料",
    "Biotechnology": "生物科技",
    "Building Products & Equipment": "建材与设备",
    "Capital Markets": "资本市场",
    "Cloud": "云计算",
    "Communication Equipment": "通信设备",
    "Computer Hardware": "计算机硬件",
    "Consumer Electronics": "消费电子",
    "Copper": "铜业",
    "Credit Services": "信贷服务",
    "Diagnostics & Research": "诊断与研究",
    "Discount Stores": "折扣零售",
    "Drug Manufacturers - General": "制药",
    "Drug Manufacturers - Specialty & Generic": "专科与仿制药",
    "Electronic Components": "电子元件",
    "Electronic Gaming & Multimedia": "电子游戏与多媒体",
    "Engineering & Construction": "工程与建筑",
    "Entertainment": "娱乐",
    "Farm & Heavy Construction Machinery": "农业与工程机械",
    "Footwear & Accessories": "鞋类与配饰",
    "Gold": "黄金",
    "HBM Storage": "HBM存储",
    "Home Improvement Retail": "家居装修零售",
    "Household & Personal Products": "家庭与个人护理用品",
    "Information Technology Services": "IT服务",
    "Insurance - Property & Casualty": "财产保险",
    "Integrated Freight & Logistics": "综合货运物流",
    "Internet Content & Information": "互联网内容与信息",
    "Internet Retail": "互联网零售",
    "Marine Shipping": "航运",
    "Medical Devices": "医疗器械",
    "Medical Instruments & Supplies": "医疗器械与耗材",
    "Oil & Gas E&P": "油气勘探与生产",
    "Oil & Gas Integrated": "综合油气",
    "Oil & Gas Midstream": "油气中游",
    "Oil & Gas Refining & Marketing": "油气炼化与销售",
    "Other Industrial Metals & Mining": "工业金属与采矿",
    "Other Precious Metals & Mining": "贵金属与采矿",
    "REIT - Healthcare Facilities": "医疗地产REIT",
    "Railroads": "铁路",
    "Restaurants": "餐饮",
    "Semiconductor Equipment & Materials": "半导体设备与材料",
    "Semiconductors": "半导体",
    "Software - Application": "应用软件",
    "Software - Infrastructure": "基础软件",
    "Solar": "太阳能",
    "Specialty Industrial Machinery": "专用工业机械",
    "Specialty Retail": "专业零售",
    "Staffing & Employment Services": "人力资源服务",
    "Telecom Services": "电信服务",
    "Tobacco": "烟草",
    "Travel Services": "旅游服务",
    "Utilities - Independent Power Producers": "独立电力生产商",
    "Utilities - Regulated Electric": "受监管电力",
}


def to_cn_name(v):
    """英文公司名→中文;已中文/无译名者原样返回。"""
    if v is None or _is_cn(v):
        return v
    return NAME_MAP.get(str(v).strip(), v)


def to_cn_industry(v):
    """英文子板块→中文;已中文/无译名者原样返回。"""
    if v is None or _is_cn(v):
        return v
    return IND_MAP.get(str(v).strip(), v)
