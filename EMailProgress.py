import poplib
import datetime
import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import re
import os
import pandas as pd


#关键字匹配
match = '走访情况记录'

#检索邮件数目，不设置为0
number = 10

#全局化参数count用来计算复合要求的邮件数
global count



# 此函数通过使用poplib实现接收邮件
def recv_email_by_pop3():
    # 要进行邮件接收的邮箱。改成自己的邮箱
    email_address = "xxxxxxxxxxxxx@xxxxxx.com"
    # 要进行邮件接收的邮箱的密码。改成自己的邮箱的密码
    # 设置 -> 账户 -> POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 -> 开启服务：POP3/SMTP服务
    # 设置 -> 账户 -> POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 -> 生成授权码
    email_password = "xxxxxxxxxxxxxx"
    # 邮箱对应的pop服务器，也可以直接是IP地址
    # 改成自己邮箱的pop服务器；qq邮箱不需要修改此值
    pop_server_host = "pop.xxxxx.com"
    # 邮箱对应的pop服务器的监听端口。改成自己邮箱的pop服务器的端口；qq邮箱不需要修改此值
    pop_server_port = 995

    #初始化count参数
    count = 0

    try:
        # 连接pop服务器。如果没有使用SSL，将POP3_SSL()改成POP3()即可其他都不需要做改动
        email_server = poplib.POP3_SSL(host=pop_server_host, port=pop_server_port, timeout=10)
        print("[pop3]----connect server success, now will check username")
    except:
        print("[pop3]----sorry the given email server address connect time out")
        exit(1)
    try:
        # 验证邮箱是否存在
        email_server.user(email_address)
        print("[pop3]----username exist, now will check password")
    except:
        print("[pop3]----sorry the given email address seem do not exist")
        exit(1)
    try:
        # 验证邮箱密码是否正确
        email_server.pass_(email_password)
        print("[pop3]----password correct,now will list email")
    except:
        print("[pop3]----sorry the given username seem do not correct")
        exit(1)

    # 邮箱中其收到的邮件的数量
    email_count = len(email_server.list()[1])

    # list()返回所有邮件的编号:
    resp, mails, octets = email_server.list()
    # 遍历所有的邮件
    # 根据设定遍历的数目number
    if number == 0:
        # 如果设定nunber=0，则不进行限制
        circletime = len(mails) + 1
    elif number > (len(mails) + 1):
        # 如果设定nunber>邮件总数，则循环数为邮件总数
        circletime = len(mails) + 1
    else:
        # 如果设定nunber<=邮件总数，则循环数为number,因为是最后一种清况，所以用else就行了，不用elif
        circletime = number
    for i in range(1, circletime):
        # 通过retr(index)读取第index封邮件的内容；这里读取最后一封，也即最新收到的那一封邮件
        resp, lines, octets = email_server.retr(i)
        # lines是邮件内容，列表形式使用join拼成一个byte变量
        email_content = b'\r\n'.join(lines)
        try:
            # 再将邮件内容由byte转成str类型
            email_content = email_content.decode('utf-8')
        except Exception as e:
            print(str(e))
            continue
        # # 将str类型转换成<class 'email.message.Message'>
        # msg = email.message_from_string(email_content)
        msg = Parser().parsestr(email_content)
        # msg.get()能获取'To'收件人，'From'发件人，'Subject'邮件主题，这里我们只需要邮件主题。
        mailname = msg.get('Subject', '')
        print("[Mailname]:",decode_str(mailname))
        # 筛选邮件
        # 正则化匹配，如果邮件名decode_str(mailname)中包含match'走访情况记录'，则符合要求
        select = re.search(match, decode_str(mailname))
        print("--Condition select:",select)
        print("--count:", count)
        if select != None:
            # 获取附件
            count = count+1
            f_list = get_att(msg, count)
            #print("f_list",f_list)


    # 关闭连接
    email_server.close()

############### 用不到！！！！！！！！！！
# indent用于缩进显示:
def parse_email(msg, indent):
    if indent == 0:
        # 邮件的From, To, Subject存在于根对象上:
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    # 需要解码Subject字符串:
                    value = decode_str(value)
                else:
                    # 需要解码Email地址:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            print('%s%s: %s' % ('  ' * indent, header, value))
    if (msg.is_multipart()):
        # 如果邮件对象是一个MIMEMultipart,
        # get_payload()返回list，包含所有的子对象:
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            # 递归打印每一个子对象:
            return parse_email(part, indent + 1)
    else:
        # 邮件对象不是一个MIMEMultipart,
        # 就根据content_type判断:
        content_type = msg.get_content_type()
        if content_type=='text/plain' or content_type=='text/html':
            # 纯文本或HTML内容:
            content = msg.get_payload(decode=True)
            # 要检测文本编码:
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            print('%sText: %s' % ('  ' * indent, content))
        else:
            # 不是文本，作为附件处理:
            print('%sAttachment: %s' % ('  ' * indent, content_type))

# 解码
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

# 猜测字符编码  用不到！！！！！！！！！！
def guess_charset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        for item in content_type.split(';'):
            item = item.strip()
            if item.startswith('charset'):
                charset = item.split('=')[1]
                break
    return charset

# 获取附件
def get_att(msg, count):
    import email
    # 初始化附件名序列
    attachment_files = []

    for part in msg.walk():
        # 获取附件名称类型
        file_name = part.get_filename()
        contType = part.get_content_type()

        # 如果有附件
        if file_name:
            # 对附件名称进行解码
            dh = email.header.decode_header(file_name)
            # 举例：dh: [(b'=?UTF-8?Q?=E8=B5=B0=E8=AE=BF=E6=83=85=E5=86=B5=E8=AE=B0=E5=BD=950305.xlsx?=', 'us-ascii')]
            # dh[0][0]是邮件名 dh[0][1]是编码类型
            filename = dh[0][0]
            if dh[0][1]:
                # 根据编码类型dh[0][1]将附件名称可读化
                filename = decode_str(str(filename, dh[0][1]))
                print("--Attachment filename:", filename)
                # 重新设置邮件名
                filename_count_path = str(count) + '.xlsx'
            # 下载附件
            data = part.get_payload(decode=True)
            # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
            att_file = open('D:\\接受邮件\\' + filename_count_path, 'wb')
            # 在attachment_files这个序列后，添加本次循环内的附件名
            attachment_files.append(filename_count_path)
            # 保存附件
            att_file.write(data)
            att_file.close()
    return attachment_files

def progress_excel():
    # 用ExcelWriter打开excel文件，引擎设置为'openpyxl'以便于保存多个sheet
    writer = pd.ExcelWriter(r'.\接受邮件\模板\走访情况记录all.xlsx', engine='openpyxl')

    # sheet1处理
    print('[Progressing sheet]:20家已签约企业...')
    # 将excel文件的第一个sheet导入dataframe
    MODEL_df = pd.DataFrame(pd.read_excel('D:\\接受邮件\\模板\\走访情况记录.xlsx', sheet_name='20家已签约企业'))
    # os.getcwd()为获取当前目录（D:\）
    # os.getcwd()+'\接受邮件'为D:\的子目录：接收邮件
    # 下一行为获取当前目录下的所有文件名，包含子目录的
    for root, dirs, files in os.walk(os.getcwd()+'\接受邮件'):
        # files是一个list，包含所有文件名
        for name in files:
            print('--Filename:',name)
            # 剔除模板excel和最终输出的excel
            if (name != '走访情况记录.xlsx')&(name != '走访情况记录all.xlsx'):
                # 设定需要处理的文件路径
                EXCELpath = 'D:\\接受邮件\\' + name
                # 打开指定excel的指定sheet
                EXCEL_df = pd.DataFrame(pd.read_excel(EXCELpath, sheet_name='20家已签约企业'))
                # 列合并
                # 由于原先dataframe中有很多NaN空元素，无法比较。所以使用fillna('')函数将这些空元素以空字符填充
                MODEL_df['客户经理'] = MODEL_df['客户经理'].fillna('') + EXCEL_df['客户经理'].fillna('')
                MODEL_df['合作平台商'+'\n'+'（如与银行签约）'] = MODEL_df['合作平台商'+'\n'+'（如与银行签约）'].fillna('') + EXCEL_df['合作平台商'+'\n'+'（如与银行签约）'].fillna('')
                MODEL_df['已签约合作事项'] = MODEL_df['已签约合作事项'].fillna('') + EXCEL_df['已签约合作事项'].fillna('')
                MODEL_df['是否已有运营商、平台商、厂家介入及内容'] = MODEL_df['是否已有运营商、平台商、厂家介入及内容'].fillna('')+EXCEL_df['是否已有运营商、平台商、厂家介入及内容'].fillna('')
                MODEL_df['潜在商机'] = MODEL_df['潜在商机'].fillna('') + EXCEL_df['潜在商机'].fillna('')
                MODEL_df['其他情况描述'] = MODEL_df['其他情况描述'].fillna('') + EXCEL_df['其他情况描述'].fillna('')
                #print(MODEL_df['合作平台商'+'\n'+'（如与银行签约）'])
    # 将处理好的dataframe转化为输出excel文件的某一个sheet
    MODEL_df.to_excel(writer, sheet_name='20家已签约企业')

    #sheet2
    print('[Progressing sheet]:100家未签约规上企业...')
    MODEL_df = pd.DataFrame(pd.read_excel('D:\\接受邮件\\模板\\走访情况记录.xlsx', sheet_name='100家未签约规上企业'))
    for root, dirs, files in os.walk(os.getcwd()+'\接受邮件'):
        for name in files:
            print('--Filename:', name)
            if (name != '走访情况记录.xlsx')&(name != '走访情况记录all.xlsx'):
                EXCELpath = 'D:\\接受邮件\\' + name
                EXCEL_df = pd.DataFrame(pd.read_excel(EXCELpath, sheet_name='100家未签约规上企业'))

                # 添加判断条件
                # MODEL_df.shape[0]为dataframe的行数
                # MODEL_df.shape[1]为dataframe的列数
                for row in range(MODEL_df.shape[0]):
                    if EXCEL_df['客户经理'].fillna('')[row] != MODEL_df['客户经理'].fillna('')[row]:
                        MODEL_df['客户经理'][row] = MODEL_df['客户经理'].fillna('')[row] + EXCEL_df['客户经理'].fillna('')[row]

                #MODEL_df['客户经理'] = MODEL_df['客户经理'].fillna('') + EXCEL_df['客户经理'].fillna('')
                MODEL_df['走访人员'] = MODEL_df['走访人员'].fillna('') + EXCEL_df['走访人员'].fillna('')
                MODEL_df['对接人层次'+'\n'+'（老板/IT/生产主管/财务/综合办公室…）'] = MODEL_df['对接人层次'+'\n'+'（老板/IT/生产主管/财务/综合办公室…）'].fillna('') + EXCEL_df['对接人层次'+'\n'+'（老板/IT/生产主管/财务/综合办公室…）'].fillna('')
                MODEL_df['沟通事项及需求'] = MODEL_df['沟通事项及需求'].fillna('') + EXCEL_df['沟通事项及需求'].fillna('')
                MODEL_df['需支撑事项'] = MODEL_df['需支撑事项'].fillna('')+EXCEL_df['需支撑事项'].fillna('')
                MODEL_df['其他情况描述'] = MODEL_df['其他情况描述'].fillna('') + EXCEL_df['其他情况描述'].fillna('')
                MODEL_df['可约见时间'] = MODEL_df['可约见时间'].fillna('') + EXCEL_df['可约见时间'].fillna('')
                #print(MODEL_df['合作平台商'+'\n'+'（如与银行签约）'])
    MODEL_df.to_excel(writer, sheet_name='100家未签约规上企业')
    
    # 保存和关闭
    writer.save()
    writer.close()


if __name__ == "__main__":
    recv_email_by_pop3()
    progress_excel()
