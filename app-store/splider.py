#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

import requests
# import urllib.request
# import re
# import xlsxwriter
import openpyxl
import json
import inquirer
# import os
import time
import os

apps = {}

def SearchAppId(app): 
  url = "http://itunes.apple.com/search?term=" + app + "&entity=software"
  r = requests.get(url)
  html = r.content
  html_doc = str(html, 'utf-8')
  data = json.loads(html_doc)
  resultCount = data['resultCount']
  results = data['results']
  print(app + ' Find ' + str(resultCount) + 'result(s)')

  for i in range(resultCount):
    app_name = results[i]['trackName']
    app_id = results[i]['trackId']
    apps[app_id] = app_name
    print('name: ' + app_name, 'id: ' + str(app_id))

def SaveContent(wb, ws, app_id, app_name, row):
    # row = 2

    for j in range(1, 11): # 只能爬取前10页
        url = 'https://itunes.apple.com/rss/customerreviews/page=' + str(j) + '/id=' + str(app_id) + '/sortby=mostrecent/json?l=en&&cc=cn'
        print('当前地址：' + url)
       
        r = requests.get(url)

        if r.status_code == 200:
          html = r.content
          html_doc = str(html, 'utf-8')
          data = json.loads(html_doc)['feed'].get('entry') or []
          for i in data:
            name = i['author']['name']['label']
            rate = i['im:rating']['label']
            user_id = i['id']['label']
            content = i['content']['label']
            version = i['im:version']['label']

            ws.cell(row=row, column=1, value=app_id)
            ws.cell(row=row, column=2, value=app_name)
            ws.cell(row=row, column=3, value=name)
            ws.cell(row=row, column=4, value=rate)
            ws.cell(row=row, column=5, value=user_id)
            ws.cell(row=row, column=6, value=content)
            ws.cell(row=row, column=7, value=version)

            row = row + 1
            print(name, rate, user_id,  content)

        # 每一页爬取延迟2秒，以防过于频繁  
        time.sleep(2)
    wb.save('app_store.xlsx')
def startFetch(wb, ws):
  
  # name_list = wb.sheetnames
  is_continue = 'yes'
  ids = []

  while is_continue == 'yes':
    app_id = input("input app's id: \n")
    
    if not int(app_id) in apps:
      break
    

    app_name = apps[int(app_id)]
    print('应用名称: ' + app_name)

    ids.append({ 'app_id': app_id, 'app_name': app_name })
    questions = [
      inquirer.List('continue',
                    message="需要继续输入appId吗",
                    choices=['yes', 'no'],
      ),
    ]
    answers = inquirer.prompt(questions)
    is_continue = answers['continue']
  
  row = ws.max_row

  for i in range(len(ids)):
    SaveContent(wb, ws, ids[i]['app_id'], ids[i]['app_name'] , row+1)

def main():

    # appid = input("请输入应用id号:")
    name = input("请输入应用名称:")
    SearchAppId(name)
    
    # if not os.path.exists(appid):
        # os.system('mkdir ' + appid)

    if os.path.exists('app_store.xlsx'):
      wb = openpyxl.load_workbook('app_store.xlsx')
      ws = wb['comment']

      startFetch(wb, ws)
    else:
      # Workbook init
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = 'comment'

      ws.cell(row=1, column=1, value="APP ID")
      ws.cell(row=1, column=2, value="APP 名称")
      ws.cell(row=1, column=3, value="昵称")
      ws.cell(row=1, column=4, value="评分")
      ws.cell(row=1, column=5, value="用户id")
      ws.cell(row=1, column=6, value="评论")
      ws.cell(row=1, column=7, value="版本")

      startFetch(wb, ws)

    
    print('Done!')

if __name__ == '__main__':
    main()
