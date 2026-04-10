from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
from io import StringIO


def scrape_production_plan():
    """
    从惠而浦OA系统抓取生产计划数据
    选择波轮工厂，获取2月和3月数据
    """

    # 配置Chrome选项
    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    # 初始化浏览器
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )

    try:
        # 1. 访问OA登录页面
        print("正在访问OA系统...")
        driver.get("https://oa.whirlpool-china.com/")
        wait = WebDriverWait(driver, 20)

        # 2. 登录
        print("正在登录...")
        try:
            username_input = wait.until(EC.presence_of_element_located((By.ID, "loginid")))
            password_input = driver.find_element(By.ID, "userpassword")
            login_btn = driver.find_element(By.ID, "login")
        except Exception:
            username_input = wait.until(EC.presence_of_element_located((By.NAME, "loginid")))
            password_input = driver.find_element(By.NAME, "userpassword")
            login_btn = driver.find_element(By.XPATH, "//input[@type='submit' or @value='登录']")

        username_input.clear()
        username_input.send_keys("HSA697")
        password_input.clear()
        password_input.send_keys("Oa654321")  # 更新密码
        login_btn.click()

        time.sleep(3)

        # 3. 访问生产计划页面
        print("正在访问生产计划页面...")
        driver.get("https://oa.whirlpool-china.com/sanyorpt/product/planview.jsp")
        time.sleep(3)

        # 4. 使用JavaScript直接设置下拉菜单值
        print("使用JavaScript选择波轮工厂...")

        wait.until(EC.presence_of_element_located((By.NAME, "orgid")))

        # 设置工厂为波轮工厂(M03)
        driver.execute_script("""
            var select = document.getElementsByName('orgid')[0];
            select.value = 'M03';
            var event = document.createEvent('HTMLEvents');
            event.initEvent('change', true, false);
            select.dispatchEvent(event);
        """)
        print("已选择波轮工厂(M03)")
        time.sleep(1)

        # 5. 设置年份和月份（2月）
        print("设置2月份...")
        driver.execute_script("""
            var yearSelect = document.getElementsByName('styear')[0];
            yearSelect.value = '2026';
            var event = document.createEvent('HTMLEvents');
            event.initEvent('change', true, false);
            yearSelect.dispatchEvent(event);
        """)

        time.sleep(0.5)

        driver.execute_script("""
            var monSelect = document.getElementsByName('stmon')[0];
            monSelect.value = '02';
            var event = document.createEvent('HTMLEvents');
            event.initEvent('change', true, false);
            monSelect.dispatchEvent(event);
        """)

        time.sleep(1)

        # 6. 点击确认按钮
        print("查询2月数据...")
        try:
            confirm_btn = driver.find_element(By.NAME, "submitbtn")
            confirm_btn.click()
        except Exception:
            driver.execute_script("document.forms[0].submit();")

        time.sleep(4)

        # 7. 提取2月数据
        print("提取2月数据...")
        table_02 = extract_table_data(driver)

        # 8. 设置3月份
        print("设置3月份...")
        driver.execute_script("""
            var monSelect = document.getElementsByName('stmon')[0];
            monSelect.value = '03';
            var event = document.createEvent('HTMLEvents');
            event.initEvent('change', true, false);
            monSelect.dispatchEvent(event);
        """)

        time.sleep(1)

        # 9. 点击确认查询3月
        print("查询3月数据...")
        try:
            confirm_btn = driver.find_element(By.NAME, "submitbtn")
            confirm_btn.click()
        except Exception:
            driver.execute_script("document.forms[0].submit();")

        time.sleep(4)

        # 10. 提取3月数据
        print("提取3月数据...")
        table_03 = extract_table_data(driver)

        # 11. 保存数据到Excel的不同Sheet
        print("保存数据...")

        # 验证数据
        if table_02 is not None and len(table_02) > 0:
            print(f"2月数据: {len(table_02)} 行")
        else:
            table_02 = None

        if table_03 is not None and len(table_03) > 0:
            print(f"3月数据: {len(table_03)} 行")
        else:
            table_03 = None

        # 保存到Excel（2月和3月分别在不同的Sheet）
        if table_02 is not None or table_03 is not None:
            save_to_excel(table_02, table_03)
            return True
        else:
            print("未获取到数据")
            return False

    except Exception as e:
        print(f"发生错误: {e}")
        import traceback
        traceback.print_exc()
        driver.save_screenshot("error_screenshot.png")
        print("已保存错误截图: error_screenshot.png")
        raise

    finally:
        input("按Enter键关闭浏览器...")
        driver.quit()


def extract_table_data(driver):
    """
    提取页面表格数据
    """
    try:
        # 等待表格加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "maintable"))
        )

        # 获取表格HTML
        table = driver.find_element(By.ID, "maintable")
        table_html = table.get_attribute('outerHTML')

        # 使用StringIO将HTML字符串转换为文件对象
        html_io = StringIO(table_html)

        # 使用pandas读取表格
        df = pd.read_html(html_io)[0]

        # 清理数据：删除完全空行
        df = df.dropna(how='all')

        print(f"提取到 {len(df)} 行数据")
        print("前5行预览:")
        print(df.head())

        return df

    except Exception as e:
        print(f"提取表格数据失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def save_to_excel(df_02, df_03):
    """
    保存数据到项目文件夹下的计划.xlsx文件
    2月数据在"2月"Sheet，3月数据在"3月"Sheet
    如果文件已存在则直接覆盖
    """
    # 获取当前脚本所在目录（项目文件夹）
    project_dir = os.path.dirname(os.path.abspath(__file__))

    # 构建文件路径
    file_path = os.path.join(project_dir, "计划.xlsx")

    # 使用ExcelWriter创建多Sheet的Excel文件
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # 写入2月数据
        if df_02 is not None and len(df_02) > 0:
            df_02.to_excel(writer, sheet_name='2月', index=False)
            print(f"2月数据已写入 '2月' Sheet ({len(df_02)} 行)")
        else:
            # 如果没有2月数据，创建一个空的Sheet或写入提示
            pd.DataFrame({'提示': ['无2月数据']}).to_excel(writer, sheet_name='2月', index=False)
            print("2月无数据")

        # 写入3月数据
        if df_03 is not None and len(df_03) > 0:
            df_03.to_excel(writer, sheet_name='3月', index=False)
            print(f"3月数据已写入 '3月' Sheet ({len(df_03)} 行)")
        else:
            # 如果没有3月数据，创建一个空的Sheet或写入提示
            pd.DataFrame({'提示': ['无3月数据']}).to_excel(writer, sheet_name='3月', index=False)
            print("3月无数据")

    print(f"数据已保存至: {file_path}")

    # 检查文件是否成功创建
    if os.path.exists(file_path):
        file_size = os.path.getsize(file_path)
        print(f"文件大小: {file_size} 字节")
    else:
        print("警告: 文件保存可能失败")


def check_dependencies():
    """
    检查必要的依赖是否已安装
    """
    try:
        import pandas
        import openpyxl
        return True
    except ImportError as e:
        print("请先安装依赖：")
        print("pip install selenium pandas openpyxl webdriver-manager")
        print(f"缺失模块: {e}")
        return False


if __name__ == "__main__":
    print("=" * 50)
    print("惠而浦生产计划数据抓取工具")
    print("=" * 50)

    if not check_dependencies():
        exit(1)

    success = scrape_production_plan()

    if success:
        print("\n数据抓取成功！")
    else:
        print("\n数据抓取失败，请检查错误信息")