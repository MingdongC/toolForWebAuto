'''
使用说明：


'''
import numpy as np
import os
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import sys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import win32api,win32con

#网站地址
addr = 'http://zhihui.guet.edu.cn/'

#答案的账号密码
ans_id = 'your id'
ans_password = 'your passward '

#答题的账号密码
qut_id = 'your id'
qut_password = 'your passward'

#文本名
text_name = 'listen_crawing.txt'

#enter_page函数进到试题界面
def enter_page(addr,id,password,browser_name):
    driver = browser_open(browser_name)
    driver.get(addr)
    driver.maximize_window()

    ##登录账号
    time.sleep(3)
    driver.find_element_by_id('name').send_keys(id)
    driver.find_element_by_id('password').send_keys(password)
    driver.find_element_by_xpath('//*[@id="Submit1"]').click()

    ##进入个人中心
    time.sleep(3)
    driver.find_element_by_link_text('进入个人中心>>').click()

    ##进到在线测试“研究生英语界面”
    time.sleep(3)
    ###跳转到新页面
    all_handles = driver.window_handles
    driver.switch_to_window(all_handles[-1])
    driver.find_element_by_link_text('在线测试').click()
    time.sleep(4)

    return all_handles, driver

#browser的启动
def browser_open(browser='chrome'):
    '''
    使用方法：打开”firefox“、”chrome“
    driver = browser_open('firefox') \ driver = browser_open('chrome')
    '''
    try:
        if browser == 'chrome':
            driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')
            return driver
        elif browser == 'firefox':
            driver = webdriver.Firefox()
            return driver
        else:
            print("Not found browser! You can enter 'firefox' or 'chrome'.")
    except Exception as msg:
        print("open browser error: %s" ,msg)

#爬取答案账号内的答案并保存到txt文件中
def ans_crawing(addr, id, password, browser_name):

    answer = {}
    ans_handles, ans_driver = enter_page(addr, id, password, browser_name)
    ans_driver.find_element_by_link_text('我的测试').click()
    time.sleep(0.2)

    ##爬取页数
    pages = ans_driver.find_elements_by_class_name('sabrosus')
    ###”去首去尾“
    pages.remove(pages[0])
    pages.remove(pages[-1])
    pages.remove(pages[-1])

    ans_page_links = []
    for i in range(len(pages)):
        ans_page_links.append(pages[i].get_attribute('href'))

    for k in range(len(ans_page_links)+1):
        exam_elems = ans_driver.find_elements_by_link_text('查看')
        for i in range(len(exam_elems)):
            exam_elems[i].click()
            time.sleep(0.2)

            ##定位到标题元素
            ans_elem_name = ans_driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/form/table/tbody/tr[1]/td/center/strong')
            ans_name = ans_elem_name.text

            ##去掉标题中的‘ ’和‘：’
            ans_name = ans_name.replace(' ', '')
            ans_name = ans_name.replace('：', '')

            ##将Analytical Listening2(2)等改成Analy. Listening2(2)
            # if ans_name.find('AnalyticalListening2(2)') != -1:
            #     ans_name = ans_name.replace('AnalyticalListening2(2)', 'Analy.Listening2(2)')
            # elif ans_name.find('AnalyticalListening2(1)') != -1:
            #     ans_name = ans_name.replace('AnalyticalListening2(1)', 'Analy.Listening2(1)')
            # elif ans_name.find('AnalyticalListening1(1)') != -1:
            #     ans_name = ans_name.replace('AnalyticalListening1(1)', 'Analy.Listening1(1)')
            # elif ans_name.find('AnalyticalListening1(2)') != -1:
            #     ans_name = ans_name.replace('AnalyticalListening1(2)', 'Analy.Listening1(2)')

            ##爬答案并进行清洗
            temp = []
            mobanhangs = ans_driver.find_elements_by_tag_name('div.mobanhang')
            for l in mobanhangs:
                    temp.append(l.text)
            ###答案清洗
            index_k = 0
            for j in range(len(temp)):
                if temp[index_k].find('您的答案') != 0:
                    temp.pop(index_k)
                else:
                    temp[index_k] = temp[index_k].replace('您的答案：','')
                    temp[index_k] = temp[index_k].replace('\n','')
                    index_k+=1
            answer[ans_name] = temp
            ans_driver.back()
            exam_elems = ans_driver.find_elements_by_link_text('查看')        #试题链接更新否则会失效
        if k < len(ans_page_links):
            ans_driver.get(ans_page_links[k])

    for l in range(len(ans_handles)):
        ans_driver.switch_to_window(ans_handles[l])
        ans_driver.close()

    ##将爬取的答案保存到txt文件中
    dict_to_txt(answer)

    return answer

#将一键多值的字典写入text
def dict_to_txt(answer):
    with open(text_name, 'w') as f:
        for key in answer:
            f.write('\n')
            f.writelines(str(key) + '：' + str(answer[key]))
        f.close()

#将text文档转成一键多值的字典导入
def txt_to_dict(text_name):
    with open(text_name, 'r') as fr:
        dict_temp = {}
        for line in fr:
            v = line.strip().split('：')
            v[1] = v[1].replace('[', '')
            v[1] = v[1].replace(']', '')
            v[1] = v[1].replace('\'', '')
            v[0] = v[0].replace(' ', '')
            v[0] = v[0].replace('：', '')
            v[0] = v[0].replace(':', '')
            temp_list = v[1].strip(', ').split(', ')
            dict_temp[v[0]] = temp_list
        fr.close()
    return dict_temp

#做题
def DoExercise(addr,id,password,answer,browser_name):

    qut_handles, qut_driver = enter_page(addr,id,password,browser_name)
    wait = WebDriverWait(qut_driver, 10)

    ##爬取页数
    pages = qut_driver.find_elements_by_class_name('sabrosus')
    ###”去首去尾“--若题的数量只有一页，则把这段屏蔽掉
    pages.remove(pages[0])
    pages.remove(pages[-1])
    pages.remove(pages[-1])

    qut_page_links = []
    for i in range(len(pages)):
        qut_page_links.append(pages[i].get_attribute('href'))

    for j in range(len(qut_page_links)+1):
        exam_elems = qut_driver.find_elements_by_link_text('查看')
        handle_pages = qut_driver.current_window_handle
        for i in range(len(exam_elems)):
            exam_elems[i].click()
            time.sleep(2)
            handles = qut_driver.window_handles
            qut_driver.switch_to_window(handles[-1])

            ##定位到标题元素
            title_elem_name = qut_driver.find_element_by_xpath('/html/body/div[2]/div[3]/form/table/tbody/tr[1]/td/center/strong')
            title_name = title_elem_name.text

            ##去掉标题中的空格
            title_name = title_name.replace(' ', '')
            title_name = title_name.replace('：', '')
            title_name = title_name.replace(':', '')

            ##将Analytical Listening2(2)等改成Analy. Listening2(2)
            if title_name.find('AnalyticalListening2(2)') != -1:
                title_name = title_name.replace('AnalyticalListening2(2)', 'Analy.Listening2(2)')
            elif title_name.find('AnalyticalListening2(1)') != -1:
                title_name = title_name.replace('AnalyticalListening2(1)', 'Analy.Listening2(1)')
            elif title_name.find('AnalyticalListening1(1)') != -1:
                title_name = title_name.replace('AnalyticalListening1(1)', 'Analy.Listening1(1)')
            elif title_name.find('AnalyticalListening1(2)') != -1:
                title_name = title_name.replace('AnalyticalListening1(2)', 'Analy.Listening1(2)')


            if title_name.find('【国际学术交流英语视听说】2') == 0:
                title_name = title_name.replace('【国际学术交流英语视听说】2','【国际学术交流英语视听说2】')

            ##判断该题是否有答案
            if title_name in answer:
                daan_list = answer[title_name]
                ##判断该题是否已经做过
                elem_tj = qut_driver.find_element_by_id('ctl00_ContentPlaceHolder1_ceshis')
                if elem_tj.text != '该测试题您无法继续提交试卷':
                    ##段落式答案填写
                    if  title_name.find('FurtherListening3') != -1 :
                        qut_driver.find_element_by_xpath('/html/body/div[2]/div[3]/form/table/tbody/tr[1]/td/table/tbody/tr[5]/td/div[2]/div[2]/span/input').click()
                        frame = qut_driver.find_element_by_xpath('/html/body/div[2]/div[3]/form/table/tbody/tr[1]/td/table/tbody/tr[5]/td/div[2]/div[2]/div/div/div[2]/iframe')
                        qut_driver.switch_to_frame(frame)
                        elem = qut_driver.find_element_by_class_name('ke-content')
                        elem.send_keys(Keys.TAB)
                        elem.send_keys(str(daan_list[0]))
                        time.sleep(1)
                        qut_driver.switch_to_default_content()
                        # print('段落式答案')

                    ##非段落式答案
                    else:
                        ###视听说1 U1: Analytical Listening2 这个题的需要重复填写3个空的答案
                        if title_name == '【国际学术交流英语视听说1】U1AnalyticalListening2':
                            fu = 3
                            for f in range(3):
                                temp = daan_list[f]
                                daan_list.insert(fu, temp)
                                fu += 1

                        daan_elems = qut_driver.find_elements_by_name('setdaan')
                        e_temp = 0
                        for e in range(len(daan_elems)):
                            if daan_elems[e_temp].get_attribute('type') == 'hidden':
                                daan_elems.remove(daan_elems[e_temp])
                            else:
                                e_temp+=1

                        for k in range(len(daan_elems)):
                            ###填空题
                            if daan_elems[k].tag_name == 'input':
                                daan_elems[k].send_keys(daan_list[k])
                                # print('做填空题')

                            ###选择题
                            elif daan_elems[k].tag_name == 'select':
                                selector = Select(daan_elems[k])
                                if daan_list[k] == 'a':
                                    selector.select_by_value('A')
                                elif daan_list[k] == 'b':
                                    selector.select_by_value('B')
                                elif daan_list[k] == 'c':
                                    selector.select_by_value('C')
                                elif daan_list[k] == 'd':
                                    selector.select_by_value('D')
                                else:
                                    print('题目选项异常：%s',daan_elems[k].text)
                                    sys.exit()
                                # print('做选择题')

                            ###未知题型
                            else:
                                print('题型不确定，请手动填写！')
                                sys.exit()

                    # check_ans = input('是否提交答案？：(y/n)')
                    # while 1 :
                    #     start = time.clock()
                    #     try:
                    #         tj_button = qut_driver.find_element_by_css_selector('html body.index_body div.content div.right_main960 form#endfrom table tbody tr td#ctl00_ContentPlaceHolder1_ceshis div#tjing input#wanchengedn.ibut')
                    #         print('已经定位到元素')
                    #         end = time.clock()
                    #         break
                    #     except:
                    #         print('还未定位到元素')
                    # print('耗时：' + str(end - start))

                    tj_button = qut_driver.find_element_by_css_selector(
                        'html body.index_body div.content div.right_main960 form#endfrom table tbody tr td#ctl00_ContentPlaceHolder1_ceshis div#tjing input#wanchengedn.ibut')

                    tj_button.click()

                    ###处理提交之后出现的“答题完成”弹框
                    text = qut_driver.switch_to_alert().text
                    time.sleep(1)
                    qut_driver.switch_to_alert().accept()

            qut_driver.close()
            qut_driver.switch_to_window(handle_pages)
            # exam_elems = qut_driver.find_elements_by_link_text('查看')        #试题链接更新否则会失效
        if j < len(qut_page_links):
            qut_driver.get(qut_page_links[j])

    for l in range(len(qut_handles)):
        qut_driver.switch_to_window(qut_handles[l])
        qut_driver.close()

    ##做完题啦！
    print("做完题啦！")
    return

#模拟鼠标移动
def hovor(driver,value):
    ActionChains(driver).move_to_element(value).perform()

#爬考试题目和内容
def exm_crawing(addr, id, password, browser_name):
    answer = {}
    ans_handles, ans_driver = enter_page(addr, id, password, browser_name)   #ans_handles是保存当前所有页面的句柄
    ans_driver.find_element_by_link_text('我的测试').click()
    time.sleep(0.2)

    ##爬取页数
    pages = ans_driver.find_elements_by_class_name('sabrosus')
    ###”去首去尾“
    pages.remove(pages[0])
    pages.remove(pages[-1])
    pages.remove(pages[-1])

    ans_page_links = []
    for i in range(len(pages)):
        ans_page_links.append(pages[i].get_attribute('href'))

    with open(text_name, 'w',encoding='utf-8') as file:

        for k in range(len(ans_page_links)+1):
            exam_elems = ans_driver.find_elements_by_link_text('查看')
            for i in range(len(exam_elems)):
                exam_elems[i].click()
                time.sleep(0.2)

                ##定位到标题元素
                ans_elem_name = ans_driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/form/table/tbody/tr[1]/td/center/strong')
                ans_name = ans_elem_name.text

                ##去掉标题中的‘ ’和‘：’
                ans_name = ans_name.replace(' ', '')
                ans_name = ans_name.replace('：', '')

                if '听力' in ans_name:
                    ##爬答案并存放
                    file.write('\n')
                    file.write(str(ans_name))
                    file.write('\n')
                    mobanlist = ans_driver.find_elements_by_tag_name('div.mobanlist')
                    for l in mobanlist:
                        templist = []
                        temp = l.text
                        templist = temp.split('\n')

                        j = 0
                        while j != len(templist):
                            if templist[j] == '2.0分' or templist[j] == '0.0分' or templist[j] == '教师讲评：' or templist[j] == '强化指南：' or templist[j] == '点击查看' :
                                templist.pop(j)

                            elif templist[j] == 'A：' or templist[j] == 'B：' or templist[j] == 'C：' or templist[j] == 'D：'or templist[j] == '您的答案：':
                                templist[j] = templist[j] + templist[j+1] + '\n'
                                templist.pop(j+1)
                                j+=1

                            else:
                                templist[j] = templist[j] + '\n'
                                j+=1


                        for b in range(len(templist)):
                            file.write(templist[b])
                        file.write('\n')

                    file.write('\n' + '='*80)
                ans_driver.back()
                exam_elems = ans_driver.find_elements_by_link_text('查看')        #试题链接更新否则会失效

            if k < len(ans_page_links):     #k是当前得页数，小于ans_page_links的长度则进入下一页
                ans_driver.get(ans_page_links[k])

        for l in range(len(ans_handles)):
            ans_driver.switch_to_window(ans_handles[l])
            ans_driver.close()
        file.close()
    return

if __name__ == "__main__":
    # read_answer = ans_crawing(addr,ans_id,ans_password,'firefox')
    # answer = txt_to_dict(text_name)
    # DoExercise(addr,qut_id,qut_password,answer,'chrome')
    exm_crawing(addr,qut_id,qut_password,'chrome')
    os.system("pause")