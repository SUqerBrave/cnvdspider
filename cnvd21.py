# coding:utf-8
import time
import re
import requests
from lxml import etree
import xlwings as xw
import sys

# Cookie 和file 需要修改

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36 Edg/92.0.902.84',
    'Cookie':'__jsluid_s=e16c9d9ddaafc6f4336d34fc55508080; __jsl_clearance_s=1630645606.195|0|W0r66g3fwax7HmnU1BRt5%2B48TFU%3D; JSESSIONID=6E4774802838A9579DA6F5150A864632',
}
file=r'D:\python\python\spider\手动cnvd'


title_list = []
dengji_list=[]
CVE_list = []
CNVD_list = []
shijian_list = []
product_list = []
miaosu_list = []
leixing_list = []
yujing_list=[]
url = 'https://www.cnvd.org.cn/flaw/list.htm'
# range(起始,结束)
start_index=0
end_index=1

for i in range(start_index, end_index):
    alurl = url
    # 起始序号+1才是页面中的序号
    print('开始从{}页开始爬取'.format(i+1))
    print('目标:   '+str(alurl))
    temp = requests.post(alurl, data={'flag': '%5BLjava.lang.String%3B%4050c0068c',
                                      # 'number': '%E8%AF%B7%E8%BE%93%E5%85%A5%E7%B2%BE%E7%A1%AE%E7%BC%96%E5%8F%B7',
                                      'numPerPage': 10, 'offset': i * 10, 'max': 10,
                                      },
                         headers=headers,
                         )

    urlhref = re.findall(r'href="/flaw/show(.*?)"', temp.text)
    print(temp.status_code)
    print(temp.text)
    print(urlhref)