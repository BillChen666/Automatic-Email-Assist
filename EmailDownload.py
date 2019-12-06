# _*_ coding: utf-8 _*_
import poplib
import email
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import zipfile
from zipfile import *
import datetime



def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    return value

def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            if header == 'Subject':
                subject = decode_str(value)
                headers['Subject'] = subject
            if header == 'From':
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            if header == 'To':
                all_cc = value.split(',')
                to = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            if header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    cc_addr = u'%s <%s>' % (name, addr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc)
    return headers

def get_email_content(message, savepath):
    attachments = []
    for part in message.walk():
        filename = part.get_filename()
        if filename:
            filename = decode_str(filename)
            data = part.get_payload(decode=True)
            abs_filename = os.path.join(savepath, filename)
            attach = open(abs_filename, 'wb')
            attachments.append(filename)
            attach.write(data)
            attach.close()
    return attachments



#解压缩rar到指定文件夹
def extractRar(zfile, path):
    # #macos 系统下
    # maccommand="sudo install -c -o $USER unrar /usr/local/bin"
    # os.system(maccommand)
    # os.system("rar x " + zfile)

    #windows 系统下
    rar_command1 = r"C:\RaR\WinRAR x "
    rar_command1+=zfile
    rar_command1+=" "
    rar_command1+=path
    if os.system(rar_command1) == 0:
        print ("Path OK.")
    else:
        print("please install UnRAR.exe")

def listdir(path, list_name):
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if (os.path.splitext(file_path)[1] == '.rar') or (os.path.splitext(file_path)[1] =='.RAR'):
            list_name.append(file_path)

if __name__ == '__main__':

    savepath = r""        #下载附件保存位置 Download path
    # 账户信息
    email = '' #account name eg. 123@123.com
    password = '' #email password
    pop3_server = '' #email box server, eg Tecent QQ mail for pop.exmail.qq.com
    # 连接到POP3服务器，带SSL的:
    server = poplib.POP3_SSL(pop3_server)
    # 可以打开或关闭调试信息:
    server.set_debuglevel(0)
    # POP3服务器的欢迎文字:
    print(server.getwelcome())
    # 身份认证:
    server.user(email)
    server.pass_(password)
    # stat()返回邮件数量和占用空间:
    msg_count, msg_size = server.stat()
    print('message count:', msg_count)
    print('message size:', msg_size, 'bytes')
    # b'+OK 237 174238271' list()响应的状态/邮件数量/邮件占用的空间大小
    resp, mails, octets = server.list()
    currentdate=""
    count=0
    foldernum=0
    for i in range(1, msg_count):
        index=msg_count-i
        try:
            resp, byte_lines, octets = server.retr(index)
        except:
            print("line too long")
            continue

        # 转码
        str_lines = []
        for x in byte_lines:
            try:
                str_lines.append(x.decode('utf-8'))
            except:
                str_lines.append(x.decode('GBK'))
        # 拼接邮件内容
        msg_content = '\n'.join(str_lines)
        # 把邮件内容解析为Message对象
        msg = Parser().parsestr(msg_content)
        headers = get_email_headers(msg)

        listname = []
        if currentdate=="":
            currentdate=headers['Date']
            folderdate=datetime.datetime.today()
            print(folderdate.strftime("%Y%m%d"))
            savepath = savepath + "\\" 
            savepath += folderdate.strftime("%Y%m%d")
            commandcreate = "mkdir "
            commandcreate += savepath
            try:
                os.system(commandcreate)
            except:
                print("ERROR, already download!")

        if headers !={}:
            if headers['Date'].find(currentdate[5:16]) == -1:
                currentdate = headers['Date']
                print(headers['Date'])
                count+=1
                     
            if count == 2:
                break
            # Add sender you want 
            if headers['From'] == " <xt.os@orientsec.com.cn>":
            #if headers['From'] == " <wenxin@orientsec.com.cn>":(test)
                attachments = get_email_content(msg, savepath)
                listdir(savepath , listname)
                print(listname)
                foldername=listname[0][:-4]
                os.system("mkdir "+foldername)
                extractRar(listname[0], foldername)
                os.system("del " + listname[0])

            #print('subject:', headers['Subject'])
            print('from:', headers['From'])
            print('to:', headers['To'])
            if 'cc' in headers:
                print('cc:', headers['Cc'])
            print('date:', headers['Date'])
            print('-----------------------------')


    server.quit()
