import requests   # 发送get/post请求，获取网站内容
import wechatsogou   # 微信公众号文章爬虫框架
import json   # json数据处理模块
import datetime   # 日期数据处理模块
import pdfkit  # 可以将文本字符串/链接/文本文件转换成为pdf
import os   # 系统文件管理
import re  # 正则匹配模块
import yagmail  # 邮件发送模块
import sys  # 项目进程管理


'''
1、从次幂数据获取公众号最新的推文链接和标题
'''
def get_data(publish_date):
    # 每日获取前一天的数据 publish_date 格式：2022-01-23
    # bid=EOdxnBO4 表示公众号 简说Python，每个公众号都有对应的bid，可以直接搜索查看
    url1 = 'https://www.cimidata.com/a/EOdxnBO4'
    r = requests.get(url1)
    # 把html的文本内容解析成html对象
    html = etree.HTML(r.text)
    # xpath 根据标签路径提取数据
    title = html.xpath('//*[@id="wrapper"]/div/div[2]/div[1]/div[2]/div/div[1]/div/div/h4/a/text()')  # 标题
    publish_time = html.xpath('//*[@id="wrapper"]/div/div[2]/div[1]/div[2]/div/div[1]/div/div/p[2]/@title')  # 发布时间
    title_url = html.xpath('//*[@id="wrapper"]/div/div[2]/div[1]/div[2]/div/div[1]/div/div/h4/a/@href')  # 文章链接
    
    # 对数据进行简单处理，选取最新发布的数据
    data = []
    for i in range(len(publish_time)):
        if publish_date in publish_time[i]:
            article = {}
            article['content_url'] = 'https://www.cimidata.com/a/EOdxnBO4' + title_url[i]
            article['title'] = title[i]
            data.append(article)
    return data

'''
2、for循环遍历，将每篇文章转化为pdf
'''
# 转化url为pdf时，调用wechatsogou中的get_article_content函数，将url中的代码提取出来转换为html字符串
# 这里先初始化一个WechatSogouAPI对象
ws_api = wechatsogou.WechatSogouAPI(captcha_break_time=3) 

def url_to_pdf(url, title, targetPath, publish_date):
    '''
    使用pdfkit生成pdf文件
    :param url: 文章url
    :param title: 文章标题
    :param targetPath: 存储pdf文件的路径
    :param publish_date: 文章发布日期，作为pdf文件名开头（标识）
    '''
    try:
        content_info = ws_api.get_article_content(url)
    except:
        return False
    # 处理后的html
    html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>{title}</title>
    </head>
    <body>
    <h2 style="text-align: center;font-weight: 400;">{title}</h2>
    {content_info['content_html']}
    </body>
    </html>
    '''
    # html字符串转换为pdf
    filename = publish_date + '-' + title
    # 部分文章标题含特殊字符，不能作为文件名
    # 去除标题中的特殊字符 win / \ : * " < > | ？ mac :  
    # 先用正则去除基本的特殊字符，python中反斜线很烦，最后用replace函数去除
    filename = re.sub('[/:*"<>|？]','',filename).replace('\\','')
    pdfkit.from_string(html, targetPath + os.path.sep + filename + '.pdf')
    return filename  # 返回存储路径，后面邮件发送附件需要

'''
3、通过邮件将新生成的文件发送到自己的邮箱
'''
def send_email(user_name, email, gzh_data):
    yag = yagmail.SMTP(user='你的发邮件的邮箱，可以和收件的是一个',password='你的POP3/SMTP服务密钥',host='smtp.163.com')
    contents = ['亲爱的 '+user_name+' 你好:<br>',
                '公众号 {0} {1}发布了{2}篇推文，推文标题分别为：<br>'.format(gzh_data['gzh_name'], gzh_data['publish_date'], len(gzh_data['save_path'])),
                '<br>'.join(gzh_data['save_path']),
                '<br>文章详细信息可以查看附件pdf内容，有问题可以在公众号%s联系作者提问。<br>'%gzh_data['gzh_name'],
                '<br><br><p align="right">公众号-%s</p>'%gzh_data['gzh_name']
                ]
    # 在邮件内容后，添加上附件路径（蛮简单实现动态添加附件，直接拼接两个列表即可哈哈哈哈）
    contents = contents + [targetPath + os.path.sep + i + '.pdf' for i in gzh_data['save_path']]
    yag.send(email, '请查看'+gzh_name+publish_date+'推文内容', contents)
    

# 程序开始
# 0、为爬取内容创建一个单独的存放目录
gzh_name = '简说Python'  # 爬取公众号名称
targetPath = os.getcwd() + os.path.sep + gzh_name
# 如果不存在目标文件夹就进行创建
if not os.path.exists(targetPath):
    os.makedirs(targetPath)
print('------pdf存储目录创建成功！')
    
# 1、从二十次幂获取微信公众号最新文章数据 
year = str(datetime.datetime.now().year)
month = str(datetime.datetime.now().month)
day = str(datetime.datetime.now().day-1)
publish_date = datetime.datetime.strptime(year+month+day,'%Y%m%d').strftime('%Y-%m-%d')  # 文章发布日期
html_data = get_data(publish_date)
if html_data:
    print('------成功获取到公众号{0}{1}推文链接！'.format(gzh_name, publish_date))
else:
    print('------公众号{0}{1}没有发布推文，请前往微信确认'.format(gzh_name, publish_date))
    sys.exit()  # 结束进程
    

# 2、for循环遍历，将每篇文章转化为pdf
save_path = []
for article in html_data:
    url = article['content_url']
    title = article['title']
    # 将文章链接内容转化为pdf，并记录存储路径，用于后面邮件发送附件
    save_path.append(url_to_pdf(url, title, targetPath, publish_date)) 
print('------pdf转换保存成功！')
    
# 3、通过邮件将新生成的文件发送到自己的邮箱
user_name = '收件人名称' # 可以写自己的名字
email = '收件邮箱地址'
gzh_data = {
    'gzh_name':gzh_name,
    'publish_date':publish_date,
    'save_path':save_path
}
send_email(user_name, email, gzh_data)
print('------邮件发送成功啦！')