import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

def get_urls(code, test_links=False):
    sina_url = f'https://finance.sina.com.cn/futures/quotes/{code}.shtml'
    shangjia_url = f'https://m.shangjia.com/qihuo/{code.lower()}/'

    if test_links:
        test_link(sina_url)
        test_link(shangjia_url)

    return sina_url, shangjia_url

def test_link(url):
    response = requests.head(url)

    if response.status_code != 200:
        print(f'Warning: {url} is not accessible. Status code: {response.status_code}')


def get_latest_price(code):
    sina_url, shangjia_url = get_urls(code)
    price = None

    # 创建一个chrome浏览器的驱动，设置为无头模式
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(options=chrome_options)

    # 尝试从新浪财经获取价格
    try:
        # 打开网页
        driver.get(sina_url)

        # 等待JavaScript加载完成
        driver.implicitly_wait(5)

        # 提取价格
        price_tag = driver.find_element(By.CSS_SELECTOR, 'td[class*="price"]')
        price = price_tag.text
        print(f'getting price succesfully from sina')
    except Exception as e:
        print(f'Error getting price from sina: {e}')

    # 如果从新浪财经获取价格失败，尝试从商家网获取价格
    if price is None or price == '--':
        try:
            # 打开网页
            driver.get(shangjia_url)

            # 等待JavaScript加载完成
            driver.implicitly_wait(10)

            # 提取价格
            price_tag = driver.find_element(By.CSS_SELECTOR, 'div[class*="remove_data"]')
            price = price_tag.text
            print(f'getting price succesfully from shangjia')
        except Exception as e:
            print(f'Error getting price from shangjia: {e}')

    # 不要忘记最后要关闭浏览器
    driver.quit()

    return price

# 使用方法
# price = get_latest_price('UR2309')
# print(f'The latest price of UR2309 is {price}')
