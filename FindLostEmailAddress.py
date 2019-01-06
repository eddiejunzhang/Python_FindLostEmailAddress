#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Edited by Eddie on 2019/1/4

from imapclient import IMAPClient
from email.header import decode_header
import re

class Imapmail(object):
 
    def __init__(self):  # 初始化数据
        self.serveraddress = None
        self.user = None
        self.passwd = None
        self.prot = None
        self.ssl = None
        self.timeout = None
        self.savepath = None
        self.server = None
 
    def client(self):  # 链接
        try:
            self.server = IMAPClient(self.serveraddress, self.prot, self.ssl, timeout=self.timeout)
            return self.server
        except BaseException as e:
            return "ERROR: >>> " + str(e)
 
    def login(self):  # 认证
        try:
            self.server.login(self.user, self.passwd)
        except BaseException as e:
            return "ERROR: >>> " + str(e)
 
    def getmaildir(self):  # 获取目录列表 [((), b'/', 'INBOX'), ((b'\\Drafts',), b'/', '草稿箱'),]
        dirlist = self.server.list_folders()
        return dirlist
 
    def getallmail(self):  # 收取所有邮件
#        print(self.server)
        self.server.select_folder('INBOX', readonly=True)  # 选择目录 readonly=True 只读,不修改,这里只选择了 收件箱
        result = self.server.search()  # 获取所有邮件总数目 [1,2,3,....]
#        '应对照文档，学习一下imapclient这个库的用法
#        print("邮件列表:", result)
        
        for _sm in result:
#        for _sm in range(2,170):
            data = self.server.fetch(_sm, ['ENVELOPE'])
#            size = self.server.fetch(_sm, ['RFC822.SIZE'])
#            print("大小", size)
            envelope = data[_sm][b'ENVELOPE']
#            print(envelope)
            try:
                subject = envelope.subject.decode()
            except:
                continue
            if subject:
                subject, de = decode_header(subject)[0]
                subject = subject if not de else subject.decode(de)
#            dates = envelope.date
#            print("主题", subject)
#            print("时间", dates)
            if subject == "系统退信":
                msgdict = self.server.fetch(_sm, ['BODY[]'])  # 获取邮件内容
                mailbody = msgdict[_sm][b'BODY[]']  # 获取邮件内容
                try:
                    mailbody1 = str(mailbody, encoding='utf-8') # 此时mailbox类型是bytes
                except UnicodeError:
                    print('UnicodeError')
                    with open(self.savepath + str(_sm) + '.eml', 'wb') as f:  # 存放邮件内容
                        f.write(mailbody)
                    continue
                except UnicodeDecodeError:
                    print('UnicodeDecodeError')
                    continue
                except UnicodeEncodeError:
                    print('UnicodeEncodeError')
                    continue
                except UnicodeTranslateError:
                    print('UnicodeTranslateError')
                    continue
                except :
                    print('Other Problem')
                    continue
                else:
                    print(mailbody.__class__)
                ''' 退信邮件中，是否是因为邮件不存在，有几种标志
                最后 4 个是对方反垃圾策略导致的
                '''
                notfoundreason=[
                        r'Mailbox not found',
                        r'User not found',
                        r'User no found',
                        r'Can not connect to',
                        r'Bad address syntax',
                        r'DNS query error',
                        r'The email account that your tried to reach does not exist',
                        r'no such user',
                        r'Recipient address rejected',
                        r'user access deny',
                        r'SpamTrap=reject mode',
                        r'relaying denied for'
                        ]
                notfoundre=[]
                notfound=[]
                notfoundreasonflag = 0
                for i in range(len(notfoundreason)):
#                    print(str(i),retext[i])
                    notfoundre=re.compile(notfoundreason[i])
                    notfound=notfoundre.findall(mailbody1)
                    if notfound:
#                        print(imap.user,str(_sm),'***retext %d 命中' % (i+1))
                        notfoundreasonflag = i+1
                        continue
#                if hitit == 0 :
#                    print('没有命中========================')
                    
                mailre=re.compile(r'[\w\d\.-]+@[\w\.-]+\.[\w\.]+')
                text=mailre.findall(mailbody1)
#                不同的邮件内容，不同位置的邮件地址就是目标邮件地址
                
                if notfoundreasonflag > 0 and notfoundreasonflag <= 8:
                    mailaddr2 = text[-2] #倒数第二个
                    mailaddr1 = text[-1] #倒数第一个
                    if mailaddr2 in mailaddr1:
                        mailaddr = mailaddr2
                    else:
                        mailaddr = mailaddr1
                    print('The target email address is: No', str(_sm), '  ',mailaddr,'===')
                    with open(self.savepath +'_list.txt','a+') as l:
                        l.write(mailaddr+','+ str(notfoundreasonflag) + '\n')
#                    print('找到邮件地址的数量 ',str(len(text)))
#                    for i in range(len(text)):
#                        print(text[i])
                    mailfile = mailaddr
                    with open(self.savepath + mailfile + '.eml', 'wb') as f:  # 存放邮件内容
                        f.write(mailbody)
                              
    def close(self):
        self.server.logout()
 
 
if __name__ == "__main__":
    for i in range(n1,n2):
        imap = Imapmail()
        imap.serveraddress = "imap.domain.com"  # 邮箱地址
        imap.user = "mailbox"+str(i)+"@service.company.cn"  # 邮箱账号
        imap.passwd = "password"  # 邮箱密码
        imap.savepath = "c:\\163mail-" + str(i) + "\\"  # 邮件存放路径
        imap.client()
        imap.login()
        imap.getallmail()
        imap.close()
