# coding:utf-8
import time
import re
import requests
from lxml import etree
import xlwings as xw
import sys

# header 和file 需要修改，headers中user-agent和cookie
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36 Edg/93.0.961.38',
    'Cookie': '__jsluid_s=e16c9d9ddaafc6f4336d34fc55508080; __jsl_clearance_s=1631260127.944|0|fc%2F%2B15aGluvaA03NFAjEQ%2Fic6d8%3D; JSESSIONID=5F3961B487128DF448A4FB76131CEF4C',
}
file = r'D:\python\python\spider\手动cnvd'
# range(起始,结束)
start_index = 11
end_index = 20
#可能爬不到最后一页，用prend记录最后一页
prend=end_index

title_list = []
dengji_list = []
CVE_list = []
CNVD_list = []
shijian_list = []
product_list = []
miaosu_list = []
leixing_list = []
yujing_list = []
url = 'https://www.cnvd.org.cn/flaw/list.htm'


for i in range(start_index, end_index):
    alurl = url
    # 起始序号+1才是页面中的序号
    print('开始从{}页开始爬取'.format(i + 1))
    print('目标:   ' + str(alurl))
    temp = requests.post(alurl, data={'flag': '%5BLjava.lang.String%3B%4050c0068c',
                                      # 'number': '%E8%AF%B7%E8%BE%93%E5%85%A5%E7%B2%BE%E7%A1%AE%E7%BC%96%E5%8F%B7',
                                      'numPerPage': 10, 'offset': i * 10, 'max': 10,
                                      },
                         headers=headers,
                         )
    allurlhref = re.findall(r'href="/flaw/show(.*?)"', temp.text)
    urlhref = []
    for each in allurlhref:
        if str(each).split('-')[1] == '2021':
            urlhref.append(each)

    print(temp.status_code)
    if temp.status_code !=200:
        print('由于' + str(resps.status_code) + '停止爬取，目前爬到了第' + str(i+1) + '页')
        prend = i + 1
        break

    print(temp.text)
    print(urlhref)
    domain = "https://www.cnvd.org.cn/flaw/show"

    for urlhrefs in urlhref:
        print('开始进行内容爬取目标   ' + str(urlhrefs))
        # 延时2秒防封禁
        time.sleep(2)
        yujing = domain + urlhrefs  # 此段url为真实的url,需要上一段代码提取cnvd的编号与domain的网端组成真正需要爬取的url
        print('yujing   ' + str(yujing))
        yujing_list.append(yujing)
        resps = requests.get(yujing, headers=headers)
        print(resps.status_code)
        if resps.status_code == 200:
            texts = resps.text
            html = etree.HTML(texts)
            title = html.xpath("//div[@class='blkContainerSblk']//h1/text()")
            print('titile   ' + str(title))
            title_list.append(title)
            CVE = re.findall(r'target="_blank">(.*?) </a><br>', texts)  # 用正则爬取
            print('CVE  ' + str(CVE))
            CVE_list.append(CVE)
            CNVD = html.xpath("normalize-space(//table[@class='gg_detail']//tr[1]/td[2]/text())")
            print('CNVD ' + str(CNVD))
            CNVD_list.append(CNVD)
            shijian = html.xpath("normalize-space(//table[@class='gg_detail']//tr[2]/td[2]/text())")
            print('shijian  ' + str(shijian))
            shijian_list.append(shijian)
            dengji = html.xpath("//td[text()=\"危害级别\"]/../td[2]/text()[2]")
            dengji_txt = str(dengji).replace('\\t', '').replace('\\n', '').replace('(', '').replace('\\r', '').replace(
                '[', '').replace(']', '').replace('\'', '').replace(', ', '').replace('(', '').replace(')', '')
            print('dengji   ' + str(dengji_txt))
            dengji_list.append(dengji_txt)
            product = html.xpath("normalize-space(//table[@class='gg_detail']//tr[4]/td[2]/text())")
            print('product  ' + str(product))
            product_list.append(product)
            miaosu = html.xpath("//td[text()=\"漏洞描述\"]/../td[2]/text()")
            miaoshu_txt = str(miaosu).replace('\\t', '').replace('\\n', '').replace('(', '').replace('\\r', '').replace(
                '[', '').replace(']', '').replace('\'', '').replace(', ', '')
            print('miaosu   ' + str(''.join(miaoshu_txt)))
            miaosu_list.append(miaoshu_txt)
            leixing = html.xpath("normalize-space(//td[text()=\"漏洞类型\"]/../td[2]/text())")
            print('leixing  ' + str(leixing))
            leixing_list.append(leixing)
        else:
            print(resps.status_code)
            print('由于'+str(resps.status_code)+'停止爬取，目前爬到了第'+str(i+1)+'页')
            prend=i+1
            break



print(title_list)
title_ex = []
for i in title_list:
    if i ==[]:
        title_ex.append('--')
    else:
        title_ex += i

print(CVE_list)
print(CNVD_list)
print(dengji_list)
print(shijian_list)
print(product_list)
print(miaosu_list)
print(leixing_list)

excelApp = xw.App(visible=True, add_book=False)
print('数据导入excel')

wb = excelApp.books.add()
wb.save(file + '\\select' + str(start_index + 1) + 'to' + str(prend + 1) + '.xlsx')

sht = wb.sheets['Sheet1']
sht.range('A1').value = 'CNVD编号'
sht.range('B1').value = '漏洞名称及类型'
sht.range('C1').value = '危害等级'
sht.range('D1').value = '影响范围'
sht.range('E1').value = '预警链接'
sht.range('F1').value = '漏洞描述'

sht.range('A2').options(transpose=True).value = CNVD_list
sht.range('B2').options(transpose=True).value = title_ex
sht.range('C2').options(transpose=True).value = dengji_list
sht.range('D2').options(transpose=True).value = product_list
sht.range('E2').options(transpose=True).value = yujing_list
sht.range('F2').options(transpose=True).value = miaosu_list

wb.save()
wb.close()
excelApp.quit()
