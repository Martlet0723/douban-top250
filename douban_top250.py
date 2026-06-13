#!/usr/bin/env python3  # 指定用 python3 解释器执行此脚本
"""
豆瓣电影 Top250 数据采集工具

功能：
  1. 自动翻页采集全部250条电影数据（10页 × 25条/页）
  2. 提取排名、名称、评分、评价人数、导演、年份、类型、链接
  3. 数据清洗（去空格、统一格式）
  4. 导出 Excel（原始数据 + 评分分布统计表）
  5. 异常处理（某页失败跳过继续，不会全崩）

用法：
  python douban_top250.py                  # 默认输出到 douban_top250.xlsx
  python douban_top250.py -o movies.xlsx   # 指定输出文件
  python douban_top250.py --delay 2.0      # 调慢请求间隔（防封IP）
"""

import argparse  # 命令行参数解析库，处理 -o、--delay 等参数
import re  # 正则表达式库，用于从 HTML 文本中提取导演、年份等信息
import sys  # 系统相关库，用于 sys.exit(1) 异常退出
import time  # 时间库，用于 time.sleep() 控制请求间隔
from typing import Optional  # 类型注解，Optional[X] 表示 X 或 None

import requests  # HTTP 请求库，发送 GET 请求获取豆瓣页面 HTML
from bs4 import BeautifulSoup, Tag  # HTML 解析库，BeautifulSoup 解析页面，Tag 表示单个标签
import pandas as pd  # 数据分析库，用于构建 DataFrame 和导出 Excel
from tqdm import tqdm  # 进度条库，显示采集进度

# ============================================================
# 配置常量
# ================== ==========================================
BASE_URL = "https://movie.douban.com/top250"  # 豆瓣 Top250 列表页基础 URL
HEADERS = {  # HTTP 请求头字典，模拟浏览器访问，避免被豆瓣反爬虫拦截
    "User-Agent": (  # 浏览器标识，告诉服务器这是 Chrome 浏览器
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "  # Windows 10 + 64位架构
        "AppleWebKit/537.36 (KHTML, like Gecko) "  # WebKit 引擎标识
        "Chrome/120.0.0.0 Safari/537.36"  # Chrome 120 版本号
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",  # 接受的响应内容类型及优先级
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",  # 优先接受中文页面
}
TIMEOUT = 15  # 单次 HTTP 请求的超时时间（秒），超过此时间未响应则抛出异常


# ============================================================
# 第1步：获取单页HTML
# ============================================================
def fetch_page(url: str) -> Optional[BeautifulSoup]:  # 函数签名：接收 URL 字符串，返回 BeautifulSoup 对象或 None
    """
    请求一个 URL，返回 BeautifulSoup 对象。

    如果请求失败（网络错误、超时、HTTP状态码非200），返回 None。
    调用方收到 None 就知道这页挂了，决定跳过还是重试。
    """
    try:  # 尝试执行可能出错的网络请求
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)  # 发送 GET 请求，传入请求头和超时时间
        resp.raise_for_status()  # 检查 HTTP 状态码，非 2xx 则抛出 HTTPError 异常
        # 豆瓣的编码有时不写在 Content-Type 里，手动设一下
        resp.encoding = "utf-8"  # 强制设置响应编码为 UTF-8，防止中文乱码
        return BeautifulSoup(resp.text, "html.parser")  # 解析 HTML 文本，返回 BeautifulSoup 对象 DOM树
    except requests.RequestException as e:  # 捕获所有 requests 库相关的异常（超时、连接错误、HTTP错误等）
        # 所有 requests 相关的异常（超时、连接错误、HTTP错误等）
        # 统一捕获，打印原因，返回 None
        tqdm.write(f"  [错误] 请求失败: {url} — {e}")  # 使用 tqdm.write 输出错误信息（避免打断进度条）
        return None  # 返回 None 表示该页请求失败

# ============================================================
# 第2步：解析单个电影条目
# ============================================================
def parse_movie_item(item: Tag) -> Optional[dict]:  # 函数签名：接收一个 BeautifulSoup Tag 对象，返回电影信息字典或 None
    """
    从一个 <li> 标签中提取电影信息。

    豆瓣每页的结构：
      <ol class="grid_view">
        <li>
          <div class="item">
            <div class="pic">    → 排名(em)、链接(a href)
            <div class="info">
              <div class="hd">   → 电影名(span.title)
              <div class="bd">
                <p class="">     → 导演、年份、类型
                <div class="star"> → 评分、评价人数

    返回格式：
      {
        "排名": int,
        "电影名称": str,
        "评分": float,
        "评价人数": int,
        "导演": str,
        "年份": int,
        "类型": str,
        "电影链接": str,
      }

    解析失败返回 None（不会因为一条数据坏了就崩掉整个程序）。
    """
    try:  # 尝试解析，任何步骤失败都会跳到 except 块
        # --- 排名 ---
        # <em class="">1</em>
        rank_tag = item.find("em")  # 查找 <em> 标签，其中包含排名数字
        rank = int(rank_tag.text.strip()) if rank_tag else None  # 如果找到标签就提取文本并转整数，否则为 None

        # --- 电影链接 ---
        # <a href="https://movie.douban.com/subject/1292052/">
        link_tag = item.find("a")  # 查找第一个 <a> 标签，href 属性指向电影详情页
        link = link_tag.get("href", "") if link_tag else ""  # 提取 href 属性的值作为电影链接

        # --- 电影名称 ---
        # 第一个 <span class="title"> 是中文名
        title_tag = item.find("span", class_="title")  # 查找 class 为 "title" 的 <span> 标签
        title = title_tag.text.strip() if title_tag else ""  # 提取文本并去除首尾空格

        # --- 导演 / 年份 / 类型 ---
        # 在 <p> 标签里（bd div下的第一个p），格式类似：
        #   导演: 弗兰克·德拉邦特 Frank Darabont&nbsp;&nbsp;&nbsp;主演: ...
        #   1994&nbsp;/&nbsp;美国&nbsp;/&nbsp;犯罪 剧情
        info_p = item.find("p")  # 查找第一个 <p> 标签，包含导演、年份、类型信息
        info_text = info_p.get_text(strip=True) if info_p else ""  # 获取纯文本内容，去除所有空白字符

        # 提取导演（":"之前的部分 + ":"之后的第一个名字段）
        # 格式：导演: 弗兰克·德拉邦特 Frank Darabont  主演: ...
        director = ""  # 初始化导演字段为空字符串
        year = None  # 初始化年份字段为 None
        genre = ""  # 初始化类型字段为空字符串

        # 导演在第一个 &nbsp; 或 "主演:" 之前
        if "导演:" in info_text or "导演：" in info_text:  # 检查文本中是否包含导演信息（中文冒号或英文冒号）
            # 取 "导演: xxx" 这一段
            match = re.search(r"导演[:：]\s*(.+?)(?:\s{2,}|主演|主\s*演)", info_text)  # 用正则匹配导演名：从"导演:"后到多个空格或"主演"前
            if match:  # 如果正则匹配成功
                director = match.group(1).strip()  # 取第一个捕获组（导演名）并去除空格

        # 年份和类型在文本后半段，格式：1994 / 美国 / 犯罪 剧情
        # 用正则提取：4位数字年份
        year_match = re.search(r"(\d{4})", info_text)  # 正则匹配4位连续数字作为年份
        if year_match:  # 如果匹配成功
            year = int(year_match.group(1))  # 提取匹配到的年份并转为整数

        # 类型在年份和国家之后：1994&nbsp;/&nbsp;美国&nbsp;/&nbsp;犯罪 剧情
        # 取最后一个 "/" 之后的部分
        parts = re.split(r"\s*/\s*", info_text)  # 用 "/" 分割文本（"/" 前后可能有空格）
        if len(parts) >= 3:  # 如果分割后至少有三段（年份 / 国家 / 类型）
            # 最后一部分就是类型（可能包含多个，空格分隔）
            genre = parts[-1].strip()  # 取最后一段作为类型并去除空格

        # --- 评分 ---
        # <span class="rating_num" property="v:average">9.7</span>
        rating_tag = item.find("span", class_="rating_num")  # 查找 class 为 "rating_num" 的 span 标签
        rating = float(rating_tag.text.strip()) if rating_tag else None  # 提取评分文本并转为浮点数

        # --- 评价人数 ---
        # 格式：<span>3293412人评价</span>
        # 遍历所有 span，找到包含"人评价"的那个
        votes = None  # 初始化评价人数为 None
        for span in item.find_all("span"):  # 遍历当前条目中所有 <span> 标签
            if span.text and "人评价" in span.text:  # 检查该 span 的文本是否包含"人评价"
                votes_match = re.search(r"(\d+)\s*人评价", span.text)  # 用正则提取"人评价"前面的数字
                if votes_match:  # 如果匹配成功
                    votes = int(votes_match.group(1))  # 提取数字并转为整数
                break  # 找到就退出循环，不继续遍历

        # 只返回完整数据
        if title and rating is not None:  # 只有在电影名和评分都不为空时才返回数据
            return {  # 返回包含电影信息的字典
                "排名": rank,  # 排名（整数）
                "电影名称": title,  # 电影中文名
                "评分": rating,  # 豆瓣评分（浮点数）
                "评价人数": votes,  # 评价人数（整数）
                "导演": director,  # 导演姓名
                "年份": year,  # 上映年份（整数）
                "类型": genre,  # 电影类型（如"剧情 喜剧"）
                "电影链接": link,  # 电影详情页 URL
            }

        return None  # 数据不完整时返回 None，调用方会过滤掉

    except Exception as e:  # 捕获解析过程中的任何异常
        tqdm.write(f"  [警告] 解析电影条目失败: {e}")  # 打印警告信息但不中断程序
        return None  # 返回 None 表示该条目解析失败


# ============================================================
# 第3步：采集全部250条数据
# ============================================================
def scrape_top250(delay: float = 1.0, dry_run: bool = False) -> list[dict]:  # delay=请求间隔秒数, dry_run=是否只采集第一页
    """
    主采集逻辑：
      1. 生成10个分页URL（start=0, 25, 50, ..., 225）
      2. 逐页请求 + 解析
      3. 某页失败 → 打印警告，跳过继续下一页
      4. 每次请求后 sleep(delay) 秒，避免触发豆瓣反爬

    返回所有成功采集的电影字典列表。
    """
    
    all_movies: list[dict] = []  # 存储所有成功采集的电影数据
    failed_pages: list[int] = []  # 记录采集失败的页码

    # 生成10页的URL列表
    urls = [f"{BASE_URL}?start={i * 25}&filter=" for i in range(10)]  # 列表推导式生成10页URL（start=0,25,50,...,225）
    if dry_run:  # 如果是试运行模式
        urls = urls[:1]  # 只看第一页（只有25条）

    print(f"\n  开始采集豆瓣电影 Top250（共 {len(urls)} 页）\n")  # 打印采集开始信息

    # tqdm 进度条包裹 urls 列表
    for i, url in enumerate(tqdm(urls, desc="  采集进度", unit="页"), start=1):  # 带进度条遍历URL列表，i从1开始
        soup = fetch_page(url)  # 调用 fetch_page 获取该页的 BeautifulSoup 对象
        if soup is None:  # 如果返回 None 表示请求失败
            failed_pages.append(i)  # 记录失败页码
            tqdm.write(f"  [跳过] 第 {i} 页采集失败，已跳过")  # 打印跳过信息
            continue  # 跳过当前页，继续下一页

        # 找到所有电影 <li> 标签
        # 豆瓣用的是 <ol class="grid_view"> 里嵌套 <li>
        movie_items = soup.find_all("div", class_="item")  # 查找所有 class 为 "item" 的 div（每个电影条目）

        if not movie_items:  # 如果该页没有找到任何电影条目
            tqdm.write(f"  [警告] 第 {i} 页未找到电影条目")  # 打印警告
            continue  # 跳过当前页

        for item in movie_items:  # 遍历该页每个电影条目 div
            movie = parse_movie_item(item)  # 调用解析函数，返回电影信息字典或 None
            if movie:  # 如果解析成功（非 None）
                all_movies.append(movie)  # 将电影数据添加到总列表

        # 礼貌性等待，别把豆瓣服务器打疼了
        if i < len(urls):  # 最后一页不用等（因为后面没有请求了）
            time.sleep(delay)  # 暂停 delay 秒后再请求下一页

    # --- 采集结果汇总 ---
    print(f"\n  采集完成：成功 {len(all_movies)} 条")  # 打印成功采集的电影总数
    if failed_pages:  # 如果有失败的页码
        print(f"  失败页: {failed_pages}（已跳过）")  # 列出失败页码

    return all_movies  # 返回所有采集到的电影数据


# ============================================================
# 第4步：数据清洗
# ============================================================
def clean_data(movies: list[dict]) -> pd.DataFrame:  # 函数签名：接收电影字典列表，返回清洗后的 DataFrame
    """
    将电影字典列表转为 DataFrame，并做清洗：
      - 按排名排序
      - 去空格
      - 补齐缺失排名（根据行号）
      - 统一类型格式
    """
    df = pd.DataFrame(movies)  # 将字典列表转换为 pandas DataFrame

    # 按排名升序排列
    if "排名" in df.columns:  # 检查 DataFrame 中是否有"排名"列
        df = df.sort_values("排名", ignore_index=True)  # 按排名升序排序，重置索引

    # 字符串列去首尾空格
    for col in ["电影名称", "导演", "类型"]:  # 遍历需要清洗的字符串列
        if col in df.columns:  # 如果该列存在于 DataFrame 中
            df[col] = df[col].astype(str).str.strip()  # 转成字符串后去除首尾空格

    # 如果某些排名为 None，用行号+1 补上
    if df["排名"].isna().any():  # 检查排名列是否有缺失值
        df["排名"] = range(1, len(df) + 1)  # 用 1 到 N 的连续序号填补缺失排名

    df["排名"] = df["排名"].astype(int)  # 确保排名列为整数类型

    return df  # 返回清洗后的 DataFrame


# ============================================================
# 第5步：导出Excel
# ============================================================
def export_excel(df: pd.DataFrame, output_path: str):  # 函数签名：接收 DataFrame 和输出文件路径
    """
    导出 Excel，包含两个 sheet：
      Sheet 1 "电影数据" — 完整的250条原始数据
      Sheet 2 "评分统计" — 评分分布统计表
    """
    # --- 评分分布统计 ---
    # 按评分区间统计电影数量
    bins = [0, 6.0, 7.0, 8.0, 8.5, 9.0, 9.5, 10.0]  # 评分区间边界值（左闭右开）
    labels = ["0-6.0", "6.0-6.9", "7.0-7.9", "8.0-8.4", "8.5-8.9", "9.0-9.4", "9.5-10.0"]  # 各区间的显示标签
    df["评分区间"] = pd.cut(df["评分"], bins=bins, labels=labels, right=False)  # 将评分分桶，right=False 表示左闭右开

    # --- 评分统计表stats ---
    # ┌───────────┬──────────┐
    # │ 评分区间   │ 电影数量  │
    # ├───────────┼──────────┤
    # │ 9.5-10.0  │    2     │
    # │ 9.0-9.4   │    3     │
    # └───────────┴──────────┘
    stats = (  # 构建评分分布统计表
        df.groupby("评分区间", observed=False)  # 按评分区间分组，observed=False 避免警告
        .size()  # 统计每组的电影数量
        .reset_index(name="电影数量")  # 重置索引并命名计数列
    )

    stats["占比"] = (stats["电影数量"] / stats["电影数量"].sum() * 100).round(1)  # 计算每个区间的百分比占比，保留1位小数
    stats["占比"] = stats["占比"].astype(str) + "%"  # 将数字转成带百分号的字符串

    # 总览行：平均分、最高分、最低分、总评价人数
    summary = pd.DataFrame(  # 创建总结统计的 DataFrame
        [
            ["平均评分", f"{df['评分'].mean():.2f}"],  # 计算平均评分并保留2位小数
            ["最高评分", f"{df['评分'].max():.2f}（{df.loc[df['评分'].idxmax(), '电影名称']}）"],  # 最高分及对应电影名
            ["最低评分", f"{df['评分'].min():.2f}（{df.loc[df['评分'].idxmin(), '电影名称']}）"],  # 最低分及对应电影名
            ["总评价人数", f"{df['评价人数'].sum():,}"],  # 评价人数的总和（带千位分隔符）
        ],
        columns=["统计项", "数值"],  # 两列：统计项名称和数值
    )

    # --- 写入 Excel ---
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:  # 使用 openpyxl 引擎创建 Excel writer（上下文管理器自动关闭）
        df.drop(columns=["评分区间"]).to_excel(writer, sheet_name="电影数据", index=False)  # 写电影数据 sheet，不写临时列和行索引
        stats.to_excel(writer, sheet_name="评分统计", index=False)  # 写评分分布统计表到第二个 sheet
        summary.to_excel(writer, sheet_name="评分统计", startcol=4, index=False)  # 总览写在评分统计 sheet 的第5列位置（不覆盖前面）

    print(f"\n  已导出: {output_path}")  # 打印导出文件路径
    print(f"    - Sheet 1: 电影数据 ({len(df)} 条)")  # 打印 Sheet 1 的行数
    print(f"    - Sheet 2: 评分统计")  # 打印 Sheet 2 名称


# ============================================================
# 第6步：命令行入口
# ============================================================
def main():  # 程序主入口函数
    parser = argparse.ArgumentParser(  # 创建命令行参数解析器
        description="豆瓣电影 Top250 数据采集工具",  # 程序描述
        formatter_class=argparse.RawDescriptionHelpFormatter,  # 保留帮助信息中的原始换行格式
        epilog="""  # 帮助信息末尾的使用示例
示例:
  python douban_top250.py                     # 默认输出 douban_top250.xlsx
  python douban_top250.py -o movies.xlsx      # 指定输出文件
  python douban_top250.py --delay 2.0         # 请求间隔2秒（慢一点，更稳）
  python douban_top250.py --no-export         # 只采集不导出，看数据预览
        """,
    )
    parser.add_argument(  # 添加 -o / --output 参数
        "-o", "--output",  # 短参数和长参数名
        default="douban_top250.xlsx",  # 默认输出文件名
        help="输出Excel文件路径 (默认: douban_top250.xlsx)",  # 帮助文本
    )
    parser.add_argument(  # 添加 --delay 参数
        "--delay",  # 参数名
        type=float,  # 参数类型为浮点数
        default=1.0,  # 默认值：每次请求间隔1秒
        help="请求间隔秒数 (默认: 1.0，建议不小于1秒)",  # 帮助文本
    )
    parser.add_argument(  # 添加 --no-export 开关
        "--no-export",  # 参数名
        action="store_true",  # 如果提供了此参数，值为 True；否则为 False
        help="只采集不导出Excel，在终端显示数据预览",  # 帮助文本
    )
    parser.add_argument(  # 添加 --dry-run 开关
        "--dry-run",  # 参数名
        action="store_true",  # 如果提供了此参数，值为 True
        help="只采集第一页（25条）测试用",  # 帮助文本
    )

    args = parser.parse_args()  # 解析命令行参数，返回包含所有参数值的命名空间

    print("=" * 60)  # 打印分隔线
    print("  豆瓣电影 Top250 数据采集工具")  # 打印工具标题
    print("=" * 60)  # 打印分隔线

    # 采集
    movies = scrape_top250(delay=args.delay, dry_run=args.dry_run)  # 调用采集函数，传入延迟和试运行参数

    if not movies:  # 如果未采集到任何电影数据
        print("\n  未采集到任何数据，程序退出。")  # 打印退出信息
        sys.exit(1)  # 以非零状态码退出，表示异常

    # 清洗
    df = clean_data(movies)  # 调用清洗函数，将电影列表转为干净的 DataFrame

    # 终端预览
    print(f"\n  数据预览:\n")  # 打印预览标题
    # 只显示关键列
    preview_cols = ["排名", "电影名称", "评分", "评价人数", "年份"]  # 定义要在预览中显示的列名列表
    available_cols = [c for c in preview_cols if c in df.columns]  # 列表推导式：从预览列中筛选 DataFrame 中实际存在的列
    print(df[available_cols].head(10).to_string(index=False))  # 打印前10行数据（不带行索引）
    print(f"\n  ... 共 {len(df)} 条")  # 打印数据总条数

    # 导出
    if not args.no_export:  # 如果没有设置 --no-export 开关（即需要导出 Excel）
        export_excel(df, args.output)  # 调用导出函数

    print("\n" + "=" * 60)  # 打印结束分隔线
    print("  完成！")  # 打印完成信息
    print("=" * 60)  # 打印结束分隔线

if __name__ == "__main__":  # Python 惯用法：当此文件作为脚本直接运行时才执行 main()
    main()  # 调用主函数启动程序