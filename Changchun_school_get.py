import requests
import json
import openpyxl


class ChangchunSchoolGet(object):
    def __init__(self):
        self.url = "https://restapi.amap.com/v3/place/text?"


    def school_search(self):
        cow = 2
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["A1"].value = "省份"
        sheet["B1"].value = "市"
        sheet["C1"].value = "邮政编码"
        sheet["D1"].value = "区"
        sheet["E1"].value = "归属地编码"
        sheet["F1"].value = "类型"
        sheet["G1"].value = "名称"
        sheet["H1"].value = "地址"
        sheet["I1"].value = "电话"
        sheet["J1"].value = "所属商圈"
        sheet["K1"].value = "经纬度"
        sheet["L1"].value = "信息刷新时间"

        for page in range(1, 20):
            print(f"正在汇总第{page}页信息；")
            parm = {
                "key": "YOUR-KEY",
                "keywords": "高等院校",
                "city": "长春",
                "citylimit": "true",
                "offset": "20",
                "page": page,
                "extensions": "all"
            }
            resp = requests.get(url=self.url, params=parm)

            json_data = json.loads(json.dumps(resp.json()))

            for poi in json_data['pois']:
                school_information_list = []

                pname = poi["pname"]  # 省份
                school_information_list.append(pname)
                cityname = poi["cityname"]  # 市
                school_information_list.append(cityname)
                postcode = poi["postcode"]  # 邮政编码
                school_information_list.append(postcode)
                adname = poi["adname"]  # 区
                school_information_list.append(adname)
                citycode = poi["citycode"]  # 归属地编码
                school_information_list.append(citycode)
                type = poi["type"]  # 类型
                school_information_list.append(type)
                name = poi["name"]  # 名称
                school_information_list.append(name)
                address = poi["address"]  # 地址
                school_information_list.append(address)
                tel = poi["tel"]  # 电话
                school_information_list.append(tel)
                business_area = poi["business_area"]  # 所属商圈
                school_information_list.append(business_area)
                location = poi["location"]  # 经纬度
                school_information_list.append(location)
                timestamp = poi["timestamp"]  # 信息刷新时间
                school_information_list.append(timestamp)

                i = -1
                for column in "ABCDEFGHIJKL":
                    cell_name = str(column) + str(cow)
                    i += 1
                    try:
                        sheet[cell_name].value = school_information_list[i]
                    except ValueError:
                        continue
                print(f"已经收录{cow-1}条信息")
                cow += 1
        wb.save("高等院校信息汇总.xlsx")


if __name__ == '__main__':
    ChangchunSchoolGet().school_search()
