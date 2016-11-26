# Methods and Parameteres for baikeEva
# author BBCK

# 调用库文件
from urllib.parse import quote, unquote
import urllib.request
from bs4 import BeautifulSoup
import re


class KeyList:

    # 词表关键字
    Aero_Avia_Keys = ['航空', '航天', '飞行器', '飞机', '导弹', '客机', '战斗机', '轰炸机', '歼击机', '攻击机', '运输机',
                      '直升机', '无人机', '火箭', '卫星', '空间站', '探测器', '飞船', '宇宙', '地球', '月球', '太阳', '深空',
                      '火星', '战机', '武器']

    Astronomy_Keys = ['天文', '天文台', '望远镜']

    Climate_Keys = ['自然', '大气', '气象', '天气', '气候', '词语概念']

    Information_Keys = ['软件', '程序', '计算机', '网络', '互联网', '协议', '链接', '代码', '编程', '网页', '通信', '通讯',
                        '传讯', '磁盘', '漏洞', '自动化', '控制']

    Energy_Keys = ['能源', '材料', '电池']


def sheet_header_insert(ws):

    head_column_no = ws.max_column

    print("表格表头共%d行，%d列" % (ws.max_row, ws.max_column))  # debug

    entry_column_index = -1

    # 判断词条名称所在列
    for item_index in range(1, (head_column_no + 1)):
        if ws.cell(row=2, column=item_index).value == "词条名称":
            entry_column_index = item_index

    print("词条名称所在列为：%d" % entry_column_index)

    # 表头写入
    # head_column_no + 1:   百科词条名
    # head_column_no + 2:   词条等级
    # head_column_no + 3:   词条网址
    # head_column_no + 4:   是否被"百度百科"收录
    # head_column_no + 5:   是否被"科普中国百科"收录
    # head_column_no + 6    是否是"特色词条"
    # head_column_no + 7:   是否为多义词
    # head_column_no + 8:   义项编码
    # head_column_no + 9:   概述字数
    # head_column_no + 10:  基本信息栏条数
    # head_column_no + 11:  一级目录条数
    # head_column_no + 12:  正文字数
    # head_column_no + 13:  参考文献条数
    # head_column_no + 14:  词条图片张数

    sheet_header = ["百科词条名",
                    "词条等级",
                    "词条网址",
                    "百度百科",
                    "科普中国",
                    "特色词条",
                    "多义词",
                    "义项编码",
                    "概述",
                    "基本信息栏",
                    "一级目录",
                    "正文",
                    "参考文献",
                    "词条图片"
                    ]

    for index in range(1, 15):
        ws.cell(row=2, column=head_column_no + index).value = sheet_header[index - 1]

    return entry_column_index


def is_kepu_baike(baike_page_soup):

    expert_icon = baike_page_soup.find("a", {"class": "posterFlag expert-icon "})
    auth_edit_icon = baike_page_soup.find("div", {"id": "authEdit"})
    auth_resource_icon = baike_page_soup.find("div", {"id": "authResource"})

    if expert_icon is not None:
        in_kepu_baike = True
    elif auth_edit_icon is not None:
        in_kepu_baike = True
    elif auth_resource_icon is not None:
        in_kepu_baike = True
    else:
        in_kepu_baike = False

    return in_kepu_baike


def is_baidu_baike(baike_page_soup):

    create_entry_tag = baike_page_soup.find("div", {"class": "create-entrance"})
    error_page_title = baike_page_soup.find("h1")

    if create_entry_tag is None and "百度百科错误页" not in error_page_title.get_text():
        return True
    else:
        return False


def get_summary_sum(baike_page_soup):

    summary_para_list = baike_page_soup.find_all("div", {"label-module": "lemmaSummary"})

    summary_words_no = 0

    if len(summary_para_list) != 0:
        for summary_para in summary_para_list:
            summary_words_no += len(summary_para.get_text())
    else:
        pass

    return summary_words_no


def get_content_sum(baike_page_soup):

    content_para_list = baike_page_soup.find_all("div", {"label-module": "para"})

    content_words = 0

    if len(content_para_list) != 0:

        for content_para in content_para_list:
            content_words += len(content_para.get_text())
    else:
        pass

    return content_words


def get_picture_sum(baike_page_soup):

    multi_album_tag = baike_page_soup.find("div", {"class": "album-list"})
    single_album_tag = baike_page_soup.find("div", {"class": "summary-pic"})

    if multi_album_tag is None:
        if single_album_tag is None:
            picture_list = baike_page_soup.find("div", {"class": "main-content"}).find_all("img")

            return len(picture_list)
        else:
            # 打开单个图册链接
            album_url = "http://baike.baidu.com" + unquote(single_album_tag.a['href'])
            album_page_soup = get_page_soup(album_url)

            return int(album_page_soup.find("div", {"class": "description"}).span.find_all("span")[1].get_text())

    else:
        # 打开多个图册链接
        album_url = "http://baike.baidu.com" + unquote(multi_album_tag.div.a['href'])
        album_page_soup = get_page_soup(album_url)

        return int(album_page_soup.find("span", {"class": "pic-num num"}).get_text())


def get_page_soup(original_url, key_word=""):

    # 生成词条网址
    search_url = original_url + key_word
    search_url = quote(search_url, safe='/:?=')
    page_response = urllib.request.urlopen(search_url)
    page_html = page_response.read()

    # 生成BeautifulSoup对象
    return BeautifulSoup(page_html, "lxml")


def get_legal_page_url(original_url, key_word):

    baidu_page_soup = get_page_soup(original_url, key_word)

    # 获取百度搜索结果信息
    result_links = baidu_page_soup.find("div", {"id": "content_left"}).find_all("h3")

    for result_link in result_links:
        # 搜索结果和关键字处理
        try:
            result_string = result_link.a.get_text().lower()
        except AttributeError:
            continue

        lower_keyword = key_word.lower()

        if lower_keyword in result_string and "百度百科" in result_string:
            print("找到合理词条")
            # 生成词条百科链接
            legal_key_word = result_link.a.get_text().split("_")[0]
            legal_baike_url = "http://baike.baidu.com/item/" + legal_key_word
            print(legal_baike_url)
            break
        else:
            continue
    else:
        legal_baike_url = None

    return legal_baike_url


def entry_evaluate(summary_sum, basic_info_sum, catalog_sum, content_sum, reference_sum, picture_sum):

    if content_sum <= 100:
        rank = 4
    elif content_sum <= 1000 or catalog_sum == 0:
        rank = 3
    elif summary_sum == 0 or basic_info_sum == 0:
        rank = 2
    elif picture_sum == 0 or reference_sum == 0:
        rank = 2
    else:
        rank = 1

    return rank


def get_legal_item_page_url(baike_page_soup, key_list):

    polysemous_tag_total = baike_page_soup.find("div", {"class": "lemmaWgt-subLemmaListTitle"})

    if polysemous_tag_total is None:
        polysemous_tag = baike_page_soup.find("div", {"class": "polysemantList-header-title"})
        total_url = polysemous_tag.find("a", text=re.compile("^共\d个义项"))["href"]
        total_url = "http://baike.baidu.com" + unquote(total_url)
        print("义项汇总页面网址：" + total_url)

        # 页面跳转
        total_page_soup = get_page_soup(total_url)
    else:
        total_page_soup = baike_page_soup

    # 获取合法义项
    item_para_list = total_page_soup.find_all("div", {"class": "para"})

    legal_item_list = []

    for item_para in item_para_list:
        for key in key_list:
            if key in item_para.a.get_text():
                legal_item_list.append(item_para)
                break
            else:
                continue

    # 判断合法义项数量
    if len(legal_item_list) == 0:
        legal_item = item_para_list[0]
    else:
        legal_item = legal_item_list[0]

    # 获取合法义项beautifulsoup对象
    item_url = "http://baike.baidu.com" + unquote(legal_item.a["href"])
    print("合理义项页面网址为：" + item_url)

    return item_url


def get_item_code(baike_page_soup):
    history_link = baike_page_soup.find("a", text="历史版本")["href"]
    print("词条历史版本链接为" + unquote(history_link))
    item_code = int(re.search(r"[^/]\d+$", history_link).group(0))
    return item_code


def excel_data_import(ws, row_now, entry_title, in_baidu_baike, in_kepu_baike, is_excellent_entry, entry_url,
                      is_polysemousentry, item_code, entry_rank=4, summary_sum=0,
                      basic_info_sum=0, catalog_sum=0, content_sum=0, reference_sum=0, picture_sum=0):

    for column_index in range(1, ws.max_column+1):
        if ws.cell(row=2, column=column_index).value == "百科词条名":
            ws.cell(row=row_now, column=column_index).value = entry_title
        elif ws.cell(row=2, column=column_index).value == "百度百科":
            ws.cell(row=row_now, column=column_index).value = in_baidu_baike
        elif ws.cell(row=2, column=column_index).value == "科普中国":
            ws.cell(row=row_now, column=column_index).value = in_kepu_baike
        elif ws.cell(row=2, column=column_index).value == "特色词条":
            ws.cell(row=row_now, column=column_index).value = is_excellent_entry
        elif ws.cell(row=2, column=column_index).value == "多义词":
            ws.cell(row=row_now, column=column_index).value = is_polysemousentry
        elif ws.cell(row=2, column=column_index).value == "义项编码":
            ws.cell(row=row_now, column=column_index).value = item_code
        elif ws.cell(row=2, column=column_index).value == "词条等级":
            ws.cell(row=row_now, column=column_index).value = entry_rank
        elif ws.cell(row=2, column=column_index).value == "词条网址":
            ws.cell(row=row_now, column=column_index).value = entry_url
        elif ws.cell(row=2, column=column_index).value == "概述":
            ws.cell(row=row_now, column=column_index).value = summary_sum
        elif ws.cell(row=2, column=column_index).value == "基本信息栏":
            ws.cell(row=row_now, column=column_index).value = basic_info_sum
        elif ws.cell(row=2, column=column_index).value == "一级目录":
            ws.cell(row=row_now, column=column_index).value = catalog_sum
        elif ws.cell(row=2, column=column_index).value == "正文":
            ws.cell(row=row_now, column=column_index).value = content_sum
        elif ws.cell(row=2, column=column_index).value == "参考文献":
            ws.cell(row=row_now, column=column_index).value = reference_sum
        elif ws.cell(row=2, column=column_index).value == "词条图片":
            ws.cell(row=row_now, column=column_index).value = picture_sum

    print("Data Import Completed")
