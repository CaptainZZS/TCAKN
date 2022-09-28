import requests
import re
import xlwt

def geturl(url):
    headers = {
        'authority': 'www.sciencedirect.com',
        'cache-control': 'max-age=0',
        'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="90", "Google Chrome";v="90"',
        'sec-ch-ua-mobile': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Mobile Safari/537.36',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'sec-fetch-site': 'none',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-user': '?1',
        'sec-fetch-dest': 'document',
        'accept-language': 'zh-CN,zh;q=0.9',
        'cookie': 'EUID=1595202f-4116-4734-af28-1f997bcbd788; sd_session_id=4fca7b3e3cbfd14eb80bb131af0c3c52bf8egxrqa; ANONRA_COOKIE=91F61FCF2E36DC0397874DA9F377A824B6EE3663E9D6D06C08EF097FB7E53417F208F7BEB090A49040BC3F662FA5E7F69F7508EE6D5F3329; has_multiple_organizations=false; id_ab=AEG; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; SD_REMOTEACCESS=eyJhY2NvdW50SWQiOiIyNDY1NzEiLCJ0aW1lc3RhbXAiOjE2MjI0MzgwNzkzNTR9; acw=4fca7b3e3cbfd14eb80bb131af0c3c52bf8egxrqa%7C%24%7CFCCE379670A529FB406F7B3A080DCA1BC8798412D048E5EE864A4BA759877224801068D362E30C757986CC0A11B1AA5636C406CDA3435FFF3FBA44D1BD4E4F2EB0469A67597464825D387A21AFA2E514; __cf_bm=60d64b027772f9dd9a5575760ae8b7d508b9220c-1622449276-1800-AdZ125Fnrx+RdrSfL6jg/sbtqe6HH08UcTawHiywGGxblMKLTFDvPmTdfZWh7ywFnsxFaeUJwMse3zPzV9lkEoU=; utt=be53-a9e6871c97156ee3605-de3196249a0055a6; fingerPrintToken=f874e0321fa48b358d7871a1fe0f9137; __gads=ID=ef36185127348cca:T=1622449278:S=ALNI_MbzFxrThcL3lQXPwhHLSI5bzM8zLg; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=-1124106680%7CMCIDTS%7C18779%7CMCMID%7C72726866164826108021450396844180152073%7CMCAAMLH-1623054080%7C11%7CMCAAMB-1623054080%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1622456480s%7CNONE%7CvVersion%7C5.2.0%7CMCAID%7CNONE%7CMCCIDH%7C-1660478464; MIAMISESSION=09ed8298-5cda-4f4b-bf48-21adaf9556e9:3799902204; mbox=session%23d8a1183d89a0406bb1104d0976537cb6%231622451265%7CPC%23111622438061086-107395.34_0%231685694205; s_pers=%20c19%3Dsd%253Abrowse%253Ajournal%253Aissue%7C1622451206697%3B%20v68%3D1622449406199%7C1622451206717%3B%20v8%3D1622449411688%7C1717057411688%3B%20v8_s%3DLess%2520than%25201%2520day%7C1622451211688%3B; s_sess=%20s_cpc%3D0%3B%20s_ppvl%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppv%3Dsd%25253Abrowse%25253Ajournal%25253Aissue%252C16%252C13%252C942%252C400%252C654%252C400%252C654%252C2%252CL%3B',
    }

    r = requests.get(url, headers=headers)
    return r.text

def pageurl(html,urllist):
    urls = re.findall(r'<a class=\"anchor article-content-title u-margin-xs-top u-margin-s-bottom\" href=\".*?\"',html)
    urlhead = 'https://www.sciencedirect.com'
    for i in urls:
        x = i.split('=')
        str1 = x[2]
        str2 = urlhead + str1[1:-1]
        urllist.append(str2)

def messageget(html,messagelist,i):
    judge = False
    tit = re.findall(r'<title>.*?</title>', html)
    if tit:
        title = tit[0].split('>')
        title = title[1][:-23]
    else:
        title = ''
    if re.findall(r'<span id="cetext[\d]*">.*?</span>', html):
        kw = re.findall(r'<span id="cetext[\d]*">.*?</span>', html)
        keyword = ''
        for i in range(len(kw)):
            str = kw[i].split('>')
            if i != len(kw) - 1:
                keyword += str[1][:-6] + ';'
            else:
                keyword += str[1][:-6]
    elif re.findall(r'<div id="kw[\d]*" class="keyword"><span>.*?</span>', html):
        kw = re.findall(r'<div id="kw[\d]*" class="keyword"><span>.*?</span>', html)
        keyword = ''
        for i in range(len(kw)):
            str = kw[i].split('>')
            if i != len(kw) - 1:
                keyword += str[2][:-6] + ';'
            else:
                keyword += str[2][:-6]
    else:
        keyword = ''
    tm = re.findall(r'\"dates\"\:{.*?}', html)
    if tm:
        time = tm[0][9:-1]
    else:
        time = ''
    if re.findall(r'Abstract</h2><div id=\"abss[\d]*\"><p id=\"spara[\d]*\">.*?</p>', html):
        abt = re.findall(r'Abstract</h2><div id=\"abss[\d]*\"><p id=\"spara[\d]*\">.*?</p>', html)
        abstract = abt[0][49:-4]
    elif re.findall(r'Abstract</h2><div id=\"abst[\d]*\"><p id=\"spar[\d]*\">.*?</p>', html):
        abt = re.findall(r'Abstract</h2><div id=\"abst[\d]*\"><p id=\"spar[\d]*\">.*?</p>', html)
        abstract = abt[0][49:-4]
    else:
        abstract = ''
    name = re.findall(r'<span class=\"text given-name\">.*?</span><span class=\"text surname\">.*?<sup>.</sup>.*?</a>', html)
    if name:
        Signature_unit = re.findall(r'\"\$\"\:{\"id\":\"staff[\d]*\"},\"_\"\:\".*?\"', html)
        if Signature_unit:
            Signature_unitdict = {}
            for i in Signature_unit:
                str1 = i.split(':')
                key = str1[3][1]
                value = str1[3][2:-1]
                Signature_unitdict[key] = value
        elif re.findall(r'\"\$\":{\"id\":\"st[\d]*\"},\"_\"\:\".*?\"', html):
            Signature_unit = re.findall(r'\"\$\":{\"id\":\"st[\d]*\"},\"_\"\:\".*?\"', html)
            Signature_unitdict = {}
            first = ord('a')
            for i in Signature_unit:
                str1 = i.split(':')
                key = chr(first)
                value = str1[3][1:-1]
                judge = True
                for i in Signature_unitdict:
                    if value == Signature_unitdict[i]:
                        judge = False
                if judge:
                    Signature_unitdict[key] = value
                    first += 1
        elif re.findall(r'{\"#name\":\"label\",\"_\":\"\w\"},{\"#name\":\"textfn\",\"_\":\".*?\"}', html):
            Signature_unit = re.findall(r'{\"#name\":\"label\",\"_\":\"\w\"},{\"#name\":\"textfn\",\"_\":\".*?\"',html)
            Signature_unitdict = {}
            first = ord('a')
            for i in Signature_unit:
                str1 = i.split(':')
                key = chr(first)
                value = str1[-1][1:-1]
                judge = True
                for i in Signature_unitdict:
                    if value == Signature_unitdict[i]:
                        judge = False
                if judge:
                    Signature_unitdict[key] = value
                    first += 1
        else:
            Signature_unitdict = ''
        namestr = ''
        for i in range(len(name)):
            str1 = name[i].split('>')
            str2 = ''
            for j in range(len(str1)):
                if '</span' in str1[j]:
                    namestr += str1[j][:-6] + ' '
                if '</sup' in str1[j]:
                    str3 = str1[j][:-5]
                    if str3 in Signature_unitdict:
                        str2 += Signature_unitdict[str3] + ';'
            if i != len(name) - 1:
                namestr += '(' + str2 + ')' + ' _' + 'and' + '_ '
            else:
                namestr += '(' + str2 + ')'

    elif re.findall(r'<span class=\"text given-name\">.*?</span><span class=\"text surname\">.*?</span>', html):
        name = re.findall(r'<span class=\"text given-name\">.*?</span><span class=\"text surname\">.*?</span>',
                          html)
        Signature_unit = []
        judge1 = 0
        namestr = ''
        if re.findall(r'\"\$\"\:{\"id\":\"staff[\d]*\"},\"_\"\:\".*?\"', html):
            Signature_unit = re.findall(r'\"\$\"\:{\"id\":\"staff[\d]*\"},\"_\"\:\".*?\"', html)
            judge1 = 1
        elif re.findall(r'\"\#name\":\"textfn\",\"_\":\".*?\"', html):
            Signature_unit = re.findall(r'\"\#name\":\"textfn\",\"_\":\".*?\"', html)
            judge1 = 2
        if len(Signature_unit) >= 2:
            if judge1 == 1:
                str1 = Signature_unit[0].split(':')
                Signature_unitdict = str1[3][1:-1]
            elif judge1 == 2:
                str1 = Signature_unit[1].split(':')
                Signature_unitdict = str1[-1][1:-1]
            for i in range(len(name)):
                str1 = name[i].split('>')
                for j in range(len(str1)):
                    if '</span' in str1[j]:
                        namestr += str1[j][:-6] + ' '
                if i != len(name) - 1:
                    namestr += '_' + 'and' + '_ '
        else:
            for i in range(len(name)):
                str1 = name[i].split('>')
                for j in range(len(str1)):
                    if '</span' in str1[j]:
                        namestr += str1[j][:-6] + ' '
            Signature_unitdict = ''
        if Signature_unitdict:
            namestr += '(' + Signature_unitdict + ')'
        else:
            namestr += '(' + 'Editor-In-Chief' + ')'
    else:
        namestr = ''
    messagelist.append([title, namestr, time, abstract, keyword])


def save_excel(urllist,messagelist):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('sheet1')
    sheet.write(0,0,'序号')
    sheet.write(0,1,'标题')
    sheet.write(0,2,'作者及署名单位')
    sheet.write(0,3,'时间')
    sheet.write(0,4,'摘要')
    sheet.write(0,5,'关键词')
    sheet.write(0,6,'url链接')
    for i in range(len(urllist)):
        sheet.write(i+1,0,i+1)
        sheet.write(i+1,1,messagelist[i][0])
        sheet.write(i+1,2,messagelist[i][1])
        sheet.write(i+1,3,messagelist[i][2])
        sheet.write(i+1,4,messagelist[i][3])
        sheet.write(i+1,5,messagelist[i][4])
        sheet.write(i+1,6,urllist[i])
    workbook.save('RCR.xls')
    print('已将数据保存至excel！')

def main():
    startpages = int(input('请输入开始的期刊号:'))
    endpages = int(input('请输入结束的期刊号:'))
    url = 'https://www.sciencedirect.com/journal/resources-conservation-and-recycling/vol/'
    urllist = []
    for i in range(startpages,endpages+1):
        url = 'https://www.sciencedirect.com/journal/resources-conservation-and-recycling/vol/'
        url += str(i)+'/suppl/C'
        html = geturl(url)
        pageurl(html,urllist)
    messagelist = []
    for i in urllist:
        print(i)
        html = geturl(i)
        messageget(html,messagelist,i)
    save_excel(urllist,messagelist)

main()
