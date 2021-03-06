# -*- coding: UTF-8 -*-

# third party lib
import redis
import requests

# python
import os
import sys
import http.client

import win32com.client
import win32com.gen_py.MSO as MSO
import win32com.gen_py.MSPPT as PO

# app
import config as conf
import constant as cons
import path as path

# 监听拉取PPT转换任务
def pull_convert_task(redis_pool):
	print('listening to redis, brpop blocking...')
	r = redis.Redis(connection_pool=redis_pool)
	keyIdBytes, pptIdBytes = r.brpop(cons.TASK_LIST_KEY, 0)

	#pptId='53da57e9db3dc4be47e74adb'
	pptId = pptIdBytes.decode()
	print('get task with pptId: '+pptId)
	return pptId

# 获取PPT文件
def fetchPptFile(pptId):
	print('fetchPptFile')
	res = requests.get('http://cloudslides.net/ppt/getPptFile?pptId='+pptId)
	pptFileData = res.content
	ppt_path = path.gen_ppt_path(pptId)
	with open(ppt_path, "wb") as ppt_file:
		ppt_file.write(pptFileData)

# 转换PPT文件
def convertPptToImage( pptId ):
	print('convertPptToImage')
	# 准备PPT应用
	Powerpoint = win32com.client.Dispatch(cons.POWERPOINT_APPLICATION_NAME);
	#Powerpoint.Visible = False
	Powerpoint.DisplayAlerts = False

	#准备路径参数
	ppt_path = path.gen_ppt_path(pptId) #PPT存放位置
	save_dir_path = path.gen_save_dir_path(pptId) #保存转换后图片的文件夹路径

    #使用PowerPoint打开本地PPT，并进行转换
	myPresentation = Powerpoint.Presentations.Open(ppt_path)
	myPresentation.SaveAs(save_dir_path, ppSaveAsJPG)
	myPresentation.Close()
    #顺手清理PPT文件
	os.remove(ppt_path);
	
	img_file_name_list = os.listdir(save_dir_path)
	pageCount = len(img_file_name_list)
	print('converted to %d pages' %(pageCount))
	return pageCount

    

# 上传幻灯图片
def uploadImages( pptId, pageCount):
	print('uploadImages')
	for index in range(1, pageCount+1):
		img_path = path.gen_single_png_path(pptId, index)
		url = "http://cloudslides.net/ppt/uploadImage?"+"pptId="+str(pptId)+"&pageId="+str(index)
		print(url)
		#上传图片
		res = requests.post(url,files={'file':open(img_path,'rb')})
		#清理图片
		os.remove(img_path);

# 发送转换完毕消息	
def sendConvertStatus(pptId, pageCount):
	print('sendConvertStatus')
	res = requests.get(conf.CLOUDSLIDES_URL+'/ppt/updateConvertStatus?pptId='+pptId+'&pageCount='+str(pageCount))
	print(str(res.status_code))	

def main():
	# 准备工作
	print('starting CloudSlides-Converter...')

	g = globals()
	for c in dir(MSO.constants): g[c] = getattr(MSO.constants, c) # globally define these
	for c in dir(PO.constants): g[c] = getattr(PO.constants, c)

	redis_pool = redis.ConnectionPool(host=conf.REDIS_IP)

	while True:
		# 监听拉取PPT转换任务
		pptId = pull_convert_task(redis_pool)

		# 获取PPT文件
		fetchPptFile(pptId)
		# 转换PPT文件
		pageCount = convertPptToImage(pptId)
		# 上传幻灯图片
		uploadImages(pptId, pageCount)
		# 发送转换完毕消息
		sendConvertStatus(pptId, pageCount)
	

if __name__ == '__main__':
    main()