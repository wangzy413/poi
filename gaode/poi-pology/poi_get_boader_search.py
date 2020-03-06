from urllib.parse import quote
from urllib import request
import json
import re

import xlrd
import xlwt
from xpinyin import Pinyin
import os
from transCoordinateSystem import gcj02_to_wgs84, gcj02_to_bd09
import area_boundary as  area_boundary
import city_grid as city_grid
import time
import collections

from requests.adapters import HTTPAdapter
import requests

poi_border_search_url = 'https://gaode.com/service/poiInfo?query_type=IDQ&pagesize=20&pagenum=1&qii=true&output=xml&cluster_state=5&need_utd=true&utd_sceneid=1000&div=PC1000&addr_poi_merge=true&is_classify=true&zoom=11&id='

def get_sheet(url):
    myWorkbook = xlrd.open_workbook(url)
    mySheets = myWorkbook.sheets()  # 获取工作表list。

    mySheet = mySheets[0]  # 通过索引顺序获取。
    return mySheet

def write_to_excel_poiborder(poi_borderslist):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet("sheet1", cell_overwrite_ok=True)

    # 第一行(列标题)
    sheet.write(0, 0, 'poiID')
    sheet.write(0, 1, 'border')

    index = 0
    if len(poi_borderslist) == 0:
        return
    for i in range(len(poi_borderslist)):
        poi_id = poi_borderslist[i][0]
        poi_border = poi_borderslist[i][1]
        sheet.write(index + 1, 0, poi_id)
        sheet.write(index + 1, 1, poi_border)

        index = index + 1

    # 最后，将以上操作保存到指定的Excel文件中
    p = Pinyin()
    data_path = os.getcwd() + os.sep + "data" + os.sep + "poi" + os.sep
    if not os.path.exists(data_path):
        os.mkdir(data_path)
    path = data_path + '边界数据' + '.xls'
    book.save(r'' + path)

    print('写入成功')
    return path


def hand(poilist, result):
    #result = json.loads(result)  # 将字符串转换为json
    pois = result['pois']
    for i in range(len(pois)):
        poilist.append(pois[i])

def search(sheet):
    poi_borders = []
    nrows = sheet.nrows
    for index in range(nrows):
        time.sleep(1)

        if index == 0:
            continue
        myCell = sheet.cell(index, 10)  # 获取单元格，i是行数，j是列数，行数和列数都是从0开始计数。
        poiID = myCell.value  # 通过单元格获取单元格数据。
        print(poiID)
        poi_border = get_boader_search(poiID)
        poi = [poiID, poi_border]
        poi_borders.append(poi)
    write_to_excel_poiborder(poi_borders)


def get_boader_search(poiID):
    req_url = poi_border_search_url + poiID
    print('请求url：', req_url)

    s = requests.Session()
    s.mount('http://', HTTPAdapter(max_retries=5))
    s.mount('https://', HTTPAdapter(max_retries=5))
    try:
        data = s.get(req_url, timeout = 5)

        seglist = re.findall(r'"name":"aoi","id":"1013","type":"text","value":"(.*?)"', data.text)
        for seg in seglist:
            print(seg)
            return seg


        return
    except requests.exceptions.RequestException as e:
        data = s.get(req_url, timeout=5)
        seglist = re.findall(r'"name":"aoi","id":"1013","type":"text","value":"(.*?)"', data.text)
        for seg in seglist:
            print(seg)

        return
    return None


if __name__ == '__main__':
    data_path = os.getcwd() + os.sep + "data" + os.sep + "poi" + os.sep + '商务住宅.xls'
    sheet = get_sheet(data_path)

    search(sheet)

