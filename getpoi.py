import json
import urllib

import xlwt
import time
import os
import pandas as pd
from datetime import datetime
from urllib import request
from urllib.parse import quote

class getpoi:
    #结果输出路径
    output_path = "E:\\PycharmProjects\\GaoDeMap\\"
    #POI分类编码表位置
    path_class = "E:\\PycharmProjects\\GaoDeMap\\amap_poicode.xlsx"
    #高德地图API_kEY
    amap_web_key = 'a5ecf00af22f2ea8cd3b0cc7b51a238b'
    poi_search_url = "https://restapi.amap.com/v3/place/text?key=%s&extensions=all&keywords=&types=%s&city=%s&citylimit=true&offset=25&page=%s&output=json"
    #要搜索城市的名称
    cityname = '沈阳'
    #要搜索POI的区县名
    areas = ['沈北新区']
    totalcontent = {}

    def __init__(self):
        data_class = self.getclass()
        for type_class in data_class:
            for area in self.areas:
                page = 1;
                if type_class['type_num'] / 10000 < 10:
                    classtype = str('0') + str(type_class['type_num'])
                else:
                    classtype = str(type_class['type_num'])
                while True:
                    if classtype[-4:] == "0000":
                        break;
                    poidata = self.get_poi(classtype, area, page);
                    poidata = json.loads(poidata)

                    if poidata['count'] == "0":
                        break;
                    else:
                        poilist = self.hand(poidata)
                        print("area：" + area + "  type：" + classtype + "  page：第" + str(page) + "页  count：" + poidata[
                            'count'] + "poilist:")
                        page += 1
                        for pois in poilist:
                            if classtype[0:2] in self.totalcontent.keys():
                                pois['bigclass'] = type_class['bigclass']
                                pois['midclass'] = type_class['midclass']
                                pois['smallclass'] = type_class['smallclass']
                                list_total = self.totalcontent[classtype[0:2]]
                                list_total.append(pois)
                            else:
                                self.totalcontent[classtype[0:2]] = []
                                pois['bigclass'] = type_class['bigclass']
                                pois['midclass'] = type_class['midclass']
                                pois['smallclass'] = type_class['smallclass']
                                self.totalcontent[classtype[0:2]].append(pois)
        for content in self.totalcontent:
            self.writeexcel(self.totalcontent[content], content)

    def writeexcel(self, data, classname):
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet = book.add_sheet(classname, cell_overwrite_ok=True)
        # 第一行(列标题)
        sheet.write(0, 0, 'x')
        sheet.write(0, 1, 'y')
        sheet.write(0, 2, 'count')
        sheet.write(0, 3, 'name')
        sheet.write(0, 4, 'adname')
        sheet.write(0, 5, 'smallclass')
        sheet.write(0, 6, 'typecode')
        sheet.write(0, 7, 'midclass')
        classname = data[0]['bigclass']
        for i in range(len(data)):
            sheet.write(i + 1, 0, data[i]['lng'])
            sheet.write(i + 1, 1, data[i]['lat'])
            sheet.write(i + 1, 2, 1)
            sheet.write(i + 1, 3, data[i]['name'])
            sheet.write(i + 1, 4, data[i]['adname'])
            sheet.write(i + 1, 5, data[i]['smallclass'])
            sheet.write(i + 1, 6, data[i]['classname'])
            sheet.write(i + 1, 7, data[i]['midclass'])
        book.save(self.output_path + self.cityname + '_' + classname + '.xls')
        print("Write Complete")

    def hand(self, poidate):
        pois = poidate['pois']
        poilist = []
        for i in range(len(pois)):
            content = {}
            content['lng'] = float(str(pois[i]['location']).split(",")[0])
            content['lat'] = float(str(pois[i]['location']).split(",")[1])
            content['name'] = pois[i]['name']
            content['adname'] = pois[i]['adname']
            content['classname'] = pois[i]['typecode']
            poilist.append(content)
        return poilist

    def readfile(self, readfilename, sheetname):
        data = pd.read_excel(readfilename, sheet_name=sheetname)
        return data

    def getclass(self):
        readcontent = self.readfile(self.path_class, "POI分类与编码（中英文）")
        data = []
        for num in range(readcontent.shape[0]):
            content = {}
            content['type_num'] = readcontent.iloc[num]['NEW_TYPE']
            content['bigclass'] = readcontent.iloc[num]['大类']
            content['midclass'] = readcontent.iloc[num]['中类']
            content['smallclass'] = readcontent.iloc[num]['小类']
            data.append(content)
        return data

    def get_poi(self, keywords, city, page):
        poiurl = self.poi_search_url % (self.amap_web_key, keywords, quote(city), page)
        data = ''
        with urllib.request.urlopen(poiurl) as f:
            data = f.read().decode('utf8')
        return data


if __name__ == "__main__":
    gp = getpoi()






