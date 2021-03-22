#coding = utf-8
from bs4 import BeautifulSoup #网页解析
import re   #正则匹配
import urllib.request,urllib.error  #指定URL，获取网页数据
import xlwt
import ssl
#全局取消证书验证（https）
ssl._create_default_https_context = ssl._create_unverified_context


def main ():
    baseUrl = "https://movie.douban.com/top250?start="
    #1.爬取网页
    datalist = getData(baseUrl)
    savePath = "douban.xls"
    #3.保存数据
    saveData(datalist,savePath)
    askUrl(baseUrl)

findLink = re.compile(r'<a href="(.*?)">')
findImgSrc = re.compile(r'<img.*src="(.*?)"',re.S)
findTitle = re.compile(r'<span class="title">(.*)</span>')
findRating = re.compile(r'<span class="rating_num".*>(.*)</span>')
findJudge = re.compile(r'<span>(\d*)人评价</span>')
findInq = re.compile(r'<span class="inq">(.*)</span>')
findBd = re.compile(r'<p class="">(.*?)</p>',re.S)   #re.S忽视换行符




def getData(baseUrl):
    datalist = []

    for i in range(0,10):   #range左闭右开
        url = baseUrl+str(i*25)
        html = askUrl(url)
        #解析内容
        soup = BeautifulSoup(html,"html.parser")
        for item in soup.find_all('div',class_="item"):   #查找符合要求的字符串行形成列表
            data = []
            item = str(item)

            link = re.findall(findLink,item)[0]
            imgSrc = re.findall(findImgSrc,item)[0]

            titles = re.findall(findTitle,item)
            if(len(titles) == 2):
                ctitle = titles[0]
                otitle = titles[1].replace("/","")
                otitle = otitle.replace("\xa0","")
                data.append(ctitle)
                data.append(otitle)
            else:
                data.append(titles[0])
                data.append(" ")  #外国名留空

            inq = re.findall(findInq,item)
            if len(inq)!=0:
                inq = inq[0].replace("。","")
                data.append(inq)
            else:
                data.append(" ")

            bd = re.findall(findBd,item)[0]
            bd = re.sub('<br(\s+)?/>(\s+)?'," ",bd) #去掉《br/》
            bd = re.sub('/'," ",bd)
            bd = bd.replace("\xa0","")
            # bd = re.sub("\xa0","",bd)
            judgeNum = re.findall(findJudge,item)[0]
            rating = re.findall(findRating,item)[0]

            data.append(link)
            data.append(imgSrc)
            data.append(bd.strip()) #去掉前后空格
            data.append(judgeNum)
            data.append(rating)

            datalist.append(data)

    # print(datalist)
    return datalist

#得到指定一个url的网页内容
def askUrl(url):
    #设置用户代理
    head = {}
    head["User-Agent"] = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"

    request = urllib.request.Request(url,headers=head)

    html = ""

    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        # print(html)
    except Exception as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)

    return html

def saveData(datalist,savePath):
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("豆瓣Top250")
    col = ("中文名","外文名","概况","电影链接","照片链接","演员信息","评价人数","评分")
    for i in range(0,8):
        sheet.write(0,i,col[i])
    for i in range(0,250):
        print("第%d条"%(i+1))
        data = datalist[i]
        for j in range(0,8):
            sheet.write(i+1,j,data[j])

    book.save(savePath)

if __name__ == '__main__':
    main()