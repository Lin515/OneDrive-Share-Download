"""
https://scer-my.sharepoint.cn/:f:/g/personal/cueion_scer_partner_onmschina_cn/EgZHUIu6XgdFtRErCg7wAXUB2kP5aBaCKT9kGOUnSwnAbQ
"""

# 如果没有安装 requests，请执行 pip install requests
import requests, re, os, os.path, json, _thread
import tkinter as tk
import tkinter.filedialog

g_ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36 Edg/92.0.902.55" # 请求头 UA 请勿修改

# 获取link下所有文件
# depth 为探索深度，若为0则探索到最深
# isReserveFolder 是否将探索到的文件夹加入列表
# 返回文件列表，文件列表每个元素为字典，每个字典固定包含以下元素：
# name 文件名
# url 文件下载url
# size 文件大小 整数，单位为字节
# path 所在文件夹
re_path = re.compile(r'\"rootFolder\":\"(.*)\",\"view\"')
re_data = re.compile(r'var\sg_listData\s=\s{\"wpq\":\"\",\"Templates\":{},\"ListData\":{\s\"Row\"\s:\s(.*"RemoteItem":\s\"\"\s*}\s*\])\s*,\"', re.S)
re_next = re.compile(r'\"NextHref\"\s:\s\"(.*)\"')
def GetFiles(link, depth=0, isReserveFolder=False) :
	print("Info: 正在解析 " + link)
	filelist = []
	try :
		response = requests.get(link, headers={"User-Agent": g_ua}, cookies=g_cookies, timeout=1.0)
		html = response.text
		# 获取当前目录
		path = re_path.search(html).group(1).replace(g_rootPath, "")
		# 载入网页中的json列表数据，具体内容格式请查看其源码
		data = json.loads(re_data.search(html).group(1))
		# 如果有下一页则继续爬取
		nextpage = re_next.search(html)
		while nextpage :
			html = requests.get(g_urlPreFolder + nextpage.group(1), headers={"User-Agent": g_ua}, cookies=g_cookies, timeout=1.0).text
			data += json.loads(re_data.search(html).group(1))
			nextpage = re_next.search(html)
	except requests.exceptions.RequestException : # requests.get 执行失败时会触发该异常
		print("Error: 请求失败 " + link)
	except AttributeError : # 正则未匹配到，会触发无 group 属性异常
		print("Error: 该Url无法获取到文件 " + link)
	else :
		for val in data :
			ele = {
				"path": path,
				"name": val["FileLeafRef"].encode("utf-8").decode("utf-8"),
			}
			if val["FSObjType"] == "0" :
				# 如果是文件
#				print("Info: 发现文件 " + path + "/" + ele["name"])
				ele["url"] = g_urlPreFile + "?UniqueId=" + val["UniqueId"][1:-1]
				ele["size"] = int(val["FileSizeDisplay"])
				filelist.append(ele)
			else :
				# 如果是文件夹
#				print("Info: 发现文件夹 " + path + "/" + ele["name"])
				ele["url"] = g_urlPreFolder + "?id=" + val["FileRef"].encode("utf-8").decode("utf-8")
				ele["size"] = -1
				if isReserveFolder :
					filelist.append(ele)
				# 继续递归解析
				if depth > 1 :
					filelist += GetFiles(ele["url"], depth-1, isReserveFolder)
				elif 0 == depth :
					filelist += GetFiles(ele["url"], 0, isReserveFolder)
	return filelist

# 下载文件
def DownloadFiles(downlist) :
	# 创建 aria2 下载列表文件
	downDir = wd_enDownloadDir.get()
	session = ""
	sessionList = []
	_file = None
	for item in downlist :
		try :
			if item["path"][0] == "/" :
				path = os.path.join(downDir, item["path"][1:])
			else :
				path = os.path.join(downDir, item["path"])
			if session != os.path.join(path, "aria2.session") :
				session = os.path.join(path, "aria2.session")
				if _file != None and not _file.closed :
					_file.close()
			if _file == None or _file.closed :
				if not os.access(path, os.F_OK) :
					os.makedirs(path)
				_file = open(session, "w")
				sessionList.append(session)
			_file.write(item["url"] + "\n")
		except IOError :
			print("Info: 下载列表文件读写错误 " + session)
	if not _file.closed :
		_file.close()

	# 调用 aria2c 完成下载
	# -s 单个文件最大线程数 -x 单个服务器最大线程数 -j 同时下载文件数 -k 最小分片大小 -c 断点续传
	# --file-allocation 文件预分配方式, 能有效降低磁盘碎片, 默认:prealloc
	# NTFS 建议使用 falloc, EXT3/4建议trunc
	print("Info: 开始进行下载")
	fedAuth = requests.utils.dict_from_cookiejar(g_cookies)["FedAuth"]
	for item in sessionList :
		command =	r'aria2c -s 16 -x 16 -j 5 -k 10M -c --file-allocation=falloc '\
					r'--header="Cookie: FedAuth={} ; User-Agent: {}" '\
					r'-d "{}" -i "{}"'.format(fedAuth, g_ua, os.path.dirname(item), item)
		os.system(command)
	# 移除临时文件
	for item in sessionList :
		os.remove(item)

# 将以字节为单位的整数 size 转换为最大以 GB 为单位的文本
def SizeString(size) :
	unit = ["B", "KB", "MB", "GB"]
	n = 0
	quo = size
	while quo / 1024 >= 1  and n < 3 : # 当 quo 除 1024 后会小于 1 时
		quo /= 1024
		n += 1
	return "{:.2f}{}".format(quo, unit[n])

# 根据 filelist 更新列表框
def UpdateFilesListBox(filelist) :
	wd_lbFiles.delete(0, "end")
	for item in filelist :
		if item["size"] >= 0 :
			wd_lbFiles.insert("end", SizeString(item["size"]) +": " + item["name"])
		elif item["size"] == -1 :
			wd_lbFiles.insert("end", "目录: " + item["name"])

# 分析链接 按钮点击
def BtAnalyse_Click() :
	global g_cookies, g_curUrl, g_rootPath, g_urlPreFolder, g_urlPreFile, g_filelist
	shareUrl = wd_enShareLink.get()
	response = requests.get(shareUrl, headers={"User-Agent": g_ua})
	g_cookies = response.cookies
	g_rootPath = re.search(r"\"rootFolder\":\"(.*)\",\"view\"", response.text).group(1)
	# 文件夹和文件的请求url前缀
	httpVDir = re.search(r"\"HttpVDir\"\s:\s\"(.*)\"\n", response.text).group(1)
	g_urlPreFolder = httpVDir + "/_layouts/15/onedrive.aspx"
	g_urlPreFile = httpVDir + "/_layouts/15/download.aspx" # ?UniqueId=
	# 获取文件列表
	g_curUrl = shareUrl
	filelist = GetFiles(shareUrl, 1, True)
	if len(filelist) == 0 :
		return
	g_filelist = filelist
	UpdateFilesListBox(g_filelist)
	wd_laRemotePath.config(text = "远程目录：/")

# 设置下载目录
def BtSetDownloadDir_Click() :
	_dir = tk.filedialog.askdirectory()
	if _dir != "" :
		wd_enDownloadDir.delete(0, "end")
		wd_enDownloadDir.insert(0, _dir.replace("/", "\\"))

# 打开选中项
def BtOpenSelect_Click() :
	# 选中多项时只取第一项
	global g_filelist, g_curUrl
	sel = wd_lbFiles.curselection()[0]
	item = g_filelist[sel]
	if item["size"] == -1 :
		filelist = GetFiles(item["url"], 1, True)
		if len(filelist) == 0 :
			return
		g_filelist = filelist
		if (g_filelist[0]["path"] != "/" and g_filelist[0]["path"] != "") :
			parent = os.path.dirname(g_filelist[0]["path"])
			g_filelist.insert(0, {
				"path": os.path.dirname(parent),
				"name": "上级目录",
				"url": g_urlPreFolder + "?id=" + g_rootPath + parent,
				"size": -1
			})
		g_curUrl = item["url"]
		UpdateFilesListBox(g_filelist)
		wd_laRemotePath.config(text = "远程目录：" + g_filelist[1]["path"])

# 列表框双击
def LbFiles_DoubleClick(event) :
	BtOpenSelect_Click()

# 下载选中项
def BtDownload_Click() :
	if not os.path.isdir(wd_enDownloadDir.get()) :
		BtSetDownloadDir_Click()
		if not os.path.isdir(wd_enDownloadDir.get()) :
			return
	cursel = wd_lbFiles.curselection()
	# 获取选中项
	downlist = []
	for idx in cursel :
		item = g_filelist[idx]
		if item["size"] == -1 :
			downlist += GetFiles(item["url"])
		else :
			downlist.append(item)
	# 创建两个线程
	try:
		_thread.start_new_thread(DownloadFiles, (downlist, ))
	except:
		print("Error: 启动下载线程失败")
		
def main() :
	global wd_window, wd_enShareLink, wd_enDownloadDir, wd_laRemotePath, wd_lbFiles

	# 创建窗口
	wd_window = tk.Tk(className=" OneDrive 分享链接爬取")
	wd_window.geometry("700x600")
	wd_window.resizable(0, 0)
	# 分享链接标签
	wd_laShareLink = tk.Label(wd_window, text="分享链接")
	wd_laShareLink.grid(row=0, column=0, padx=5, pady=5)
	# 分享链接文本框
	wd_enShareLink = tk.Entry(wd_window, width=75)
	wd_enShareLink.grid(row=0, column=1, padx=5, pady=5)
	# 下载目录标签
	wd_laDownloadDir = tk.Label(wd_window, text="下载目录")
	wd_laDownloadDir.grid(row=1, column=0, padx=5, pady=5)
	# 下载目录编文本框
	wd_enDownloadDir = tk.Entry(wd_window, width=75)
	wd_enDownloadDir.grid(row=1, column=1, padx=5, pady=5)
	# 远程目录标签
	wd_laRemotePath = tk.Label(wd_window, text="远程目录：无")
	wd_laRemotePath.grid(row=2, column=0, columnspan=2, padx=5, sticky=tk.W+tk.S)
	# 文件列表
	wd_lbFiles = tk.Listbox(wd_window, height=27, width=40,	selectmode=tk.EXTENDED)
	wd_lbFiles.bind("<Double-Button-1>", LbFiles_DoubleClick) # 鼠标双击事件
	wd_lbFiles.grid(row=3, rowspan=50, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E)

	# 分析按钮
	wd_btAnalyse = tk.Button(wd_window, width=10, text="分析链接", command=BtAnalyse_Click)
	wd_btAnalyse.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W+tk.E)
	# 设置下载目录按钮
	wd_btSetDownloadDir = tk.Button(wd_window, text="设置下载目录", command=BtSetDownloadDir_Click)
	wd_btSetDownloadDir.grid(row=1, column=2, padx=5, pady=5, sticky=tk.N+tk.W+tk.E)
	# 打开按钮
	wd_btOpenSelect = tk.Button(wd_window, text="打开", command=BtOpenSelect_Click)
	wd_btOpenSelect.grid(row=3, column=2, padx=5, pady=5, sticky=tk.N+tk.W+tk.E)
	# 下载按钮
	wd_btDownload = tk.Button(wd_window, text="下载选中", command=BtDownload_Click)
	wd_btDownload.grid(row=4, column=2, padx=5, pady=5, sticky=tk.N+tk.W+tk.E)
	wd_window.mainloop()

if __name__ == "__main__" :
	main()
