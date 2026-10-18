"""抓取豆瓣电影 Top 250，并将电影信息保存为 Excel 文件。"""

from pathlib import Path

import requests
import xlwt
from bs4 import BeautifulSoup


# 请求网页时使用的浏览器标识，降低被网站识别为程序请求的概率。
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Mobile Safari/537.36 Edg/131.0.0.0"
    )
}
PAGE_SIZE = 25  # 豆瓣 Top 250 每页展示的电影数量。
PAGE_COUNT = 10  # 共抓取 10 页，合计 250 部电影。
OUTPUT_FILE = Path(__file__).with_name("豆瓣电影Top250.xls")  # Excel 保存位置：脚本所在目录。
HEADERS_ROW = ("排名", "电影名称", "海报链接", "评分", "导演/演员/年份等", "短评")  # Excel 表头。

# requests.get() 方法会在网络错误时抛出异常，例如连接超时、DNS 解析失败等。
def request_get(url: str) -> str:
    """请求指定网页；网络错误或 HTTP 错误会直接抛出，避免静默失败。"""
    """静默失败的意思是指程序在遇到错误时没有给出任何提示或反馈，而是悄无声息地继续运行，可能导致后续操作出现问题而不易察觉。"""
    # requests.get() 方法会在网络错误时抛出异常，例如连接超时、DNS 解析失败等。
    response = requests.get(url, headers=HEADERS, timeout=15)
    # 如果响应状态码不是 200，raise_for_status() 会抛出 HTTPError 异常，避免静默失败。
    response.raise_for_status()
    # requests 库会根据响应头自动解码内容，response.text 会返回解码后的字符串。
    return response.text


def get_movies(page: int) -> list[tuple[str, str, str, str, str, str]]:
    """抓取并解析指定页的电影，返回每部电影的字段数据。

    page 从 0 开始计数：0 表示第一页，1 表示第二页，以此类推。
    """
    # start 参数是当前页之前已展示的电影数量。
    url = f"https://movie.douban.com/top250?start={page * PAGE_SIZE}&filter="
    soup = BeautifulSoup(request_get(url), "lxml")
    movie_items = soup.select("ol.grid_view > li")
    if not movie_items:
        raise RuntimeError(f"第 {page + 1} 页没有找到电影条目，页面可能被拦截或结构已变化。")

    movies: list[tuple[str, str, str, str, str, str]] = []
    for item in movie_items:
        # 从当前电影条目中查找需要保存的信息。
        title_node = item.select_one("span.title")
        image_node = item.select_one(".pic img")
        rank_node = item.select_one(".pic em")
        score_node = item.select_one("span.rating_num")
        details_node = item.select_one("div.bd p")

        # 这些字段是生成电影记录所必需的，缺失时给出明确错误。
        missing = [
            label
            for label, node in (
                ("电影名称", title_node),
                ("海报", image_node),
                ("排名", rank_node),
                ("评分", score_node),
                ("电影信息", details_node),
            )
            if node is None
        ]
        if missing:
            raise ValueError(f"电影条目缺少必要字段：{', '.join(missing)}")

        # 上面的完整性检查已确认这些节点存在，显式断言便于类型检查器识别。
        assert title_node is not None
        assert image_node is not None
        assert rank_node is not None
        assert score_node is not None
        assert details_node is not None

        # 短评并非每部电影都有，因此缺失时保存为空字符串。
        quote_node = item.select_one("span.inq")
        image_url = image_node.get("src", "")
        if not isinstance(image_url, str):
            image_url = ""

        movie: tuple[str, str, str, str, str, str] = (
            # 按 Excel 表头顺序保存：排名、名称、海报、评分、详情、短评。
            rank_node.get_text(strip=True),
            title_node.get_text(strip=True),
            image_url,
            score_node.get_text(strip=True),
            " ".join(details_node.stripped_strings),
            quote_node.get_text(strip=True) if quote_node else "",
        )
        movies.append(movie)

    return movies


def save_to_excel(movies: list[tuple[str, str, str, str, str, str]]) -> None:
    """将电影数据写入 Excel 工作簿并保存到配置的输出路径。"""
    workbook = xlwt.Workbook(encoding="utf-8")
    sheet = workbook.add_sheet("豆瓣电影Top250")

    # 第一行写入列标题；电影数据从第二行开始。
    for column, title in enumerate(HEADERS_ROW):
        sheet.write(0, column, title)

    for row, movie in enumerate(movies, start=1):
        for column, value in enumerate(movie):
            sheet.write(row, column, value)

    workbook.save(str(OUTPUT_FILE))


def main() -> None:
    """依次抓取所有页面，合并电影数据并导出 Excel 文件。"""
    movies: list[tuple[str, str, str, str, str, str]] = []
    for page in range(PAGE_COUNT):
        page_movies = get_movies(page)
        movies.extend(page_movies)
        print(f"已抓取第 {page + 1}/{PAGE_COUNT} 页：{len(page_movies)} 部电影")

    save_to_excel(movies)
    print(f"抓取完成，共 {len(movies)} 部电影，已保存到：{OUTPUT_FILE}")


if __name__ == "__main__":
    # 仅在直接运行本文件时开始抓取；被其他模块导入时不自动发起请求。
    main()
