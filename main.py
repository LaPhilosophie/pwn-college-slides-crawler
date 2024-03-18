from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import yaml


# 打开并读取文件
with open('dojo.yml', 'r') as file:
    lines = file.readlines()

# 用于保存提取出来的URLs的列表
urls = []

# 前缀
prefix = "https://docs.google.com/presentation/d/"

# 遍历每一行进行检查
for line in lines:
    # 如果行包含'slides:'
    if 'slides:' in line:
        # 提取'slides:'后的字符串，并去除前后的空格和换行符
        slide_id = line.split('slides:')[-1].strip()
        # 拼接URL并添加到列表
        urls.append(prefix + slide_id)

# 创建Chrome浏览器实例
driver = webdriver.Chrome()

# 对于提取出来的每个URL，进行浏览器操作
for url in urls:
    driver.get(url)
    sleep(5)  # 等待页面加载
    driver.find_element(By.CSS_SELECTOR, "#docs-file-menu").click()

    # 以下代码可能需要根据实际页面元素进行调整
    for each in driver.find_elements(By.CSS_SELECTOR, "div.goog-menuitem-content"):  # 所有的包含多个菜单项的下拉菜单
        if "下载" in each.text:  # 检查下拉菜单的文本中是否包含"下载"这个词
            each.click()  # 如果包含，点击这个下拉菜单
            break  # 假设只有一个下载菜单，点击后就可以退出循环

    # 点击下载为.pptx格式的选项
    sleep(2)  # 等待下载菜单加载
    for each in driver.find_elements(By.CSS_SELECTOR, "div.goog-menuitem-content"):
        if ".pptx" in each.text:
            each.click()
            break  # 假设只有一个.pptx选项，点击后就可以退出循环

    sleep(5)  # 等待下载完成

# 关闭浏览器
driver.quit()
