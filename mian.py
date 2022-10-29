import os
import signal
from tkinter import filedialog, messagebox

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import jissa_file


def login():
    driver.get(os.getenv("URL"))

    user_name = driver.find_element(By.NAME, "username")
    user_name.send_keys(os.getenv("USER_NAME"))

    password = driver.find_element(By.NAME, "password")
    password.send_keys(os.getenv("PASSWORD"))

    loggin = driver.find_element(By.XPATH, "/html/body/form/input[2]")
    loggin.click()


def change_to_input_frame():
    """
    pi-netの画面はいくつかのフレームに分かれているので適宜切り替える
    """

    # 画面左のフレームに切替
    iframe1 = driver.find_element(By.XPATH, "/html/frameset/frame[1]")
    driver.switch_to.frame(iframe1)

    iframe2 = driver.find_element(By.XPATH, "/html/frameset/frame[2]")
    driver.switch_to.frame(iframe2)

    # 支給品在庫報告を表示
    driver.execute_script("javascript:return showPSP()")

    # 中央フレームに切替
    driver.switch_to.default_content()
    iframe3 = driver.find_element(By.XPATH, "/html/frameset/frame[2]")
    driver.switch_to.frame(iframe3)

    # 工区入力
    kouku = driver.find_element(By.NAME, "KEYKUK")
    kouku.send_keys("18A")

    # ステータスを在庫未報告にする
    status = Select(driver.find_element(By.NAME, "HYOUJI"))
    status.select_by_value("V")

    # 支給区分を切り替える
    pay_category = Select(driver.find_element(By.NAME, "KEYUMU"))
    if jissa.is_paid:
        pay_category.select_by_value("J")
    else:
        pay_category.select_by_value("L")


def display_update():
    """
    画面右上の検索/更新ボタンを実行する
    """

    driver.execute_script("javascript:clickButton(''); return top.frSubMenu.clickExecute();")


def input_inventory_quantity() -> bool:
    """
    pinetの対象図番を実査登録確認リストから取得し、実査数をpinetに入力する。
    図番がなくるもしくは、99ページに到達したらFalseを返す。
    それ以外はTrueを返す。
    :return: bool
    """
    parts_no = ""
    for row in range(2, 16):
        parts_no = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[2]").text.strip()

        if parts_no == "":
            break

        inventory_quantity: float = jissa.get_inventory_quantity(parts_no)
        if inventory_quantity is not None:
            month_end_stock = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[8]/input")
            month_end_stock.clear()
            month_end_stock.send_keys(inventory_quantity)

            checkbox = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[1]/input[2]")
            driver.execute_script("arguments[0].click();", checkbox)
        else:
            not_found_zuban_list.append(parts_no)

    # 在庫報告ボタン
    driver.execute_script("javascript:top.frSubMenu.submitOnXXAID('5');")

    # 99ページまでしか表示されないので到達したらFalseを返す
    page_display = driver.find_element(By.XPATH, '/html/body/form/table[2]/tbody/tr/td[1]')
    if page_display.text == "PAPSP 99":
        return False
    else:
        return parts_no != ""


if __name__ == "__main__":
    load_dotenv()

    not_found_zuban_list = []

    file_type = [("Excel File", "*.xlsx")]
    base_dir = os.getenv("INIT_DIR")
    file_path = filedialog.askopenfilename(filetypes=file_type, initialdir=base_dir)

    if file_path != "":
        jissa = jissa_file.JissaFile(file_path)

        driver = webdriver.Edge()
        driver.implicitly_wait(10)

        login()
        change_to_input_frame()
        display_update()
        while input_inventory_quantity():  # 1ページごと入力して報告
            pass
            # テスト用(在庫報告ボタンを押すと次ページが表示されるので不要)
            # driver.execute_script("javascript:return top.frSubMenu.clickNextPage();")

        os.kill(driver.service.process.pid, signal.SIGTERM)
        messagebox.showinfo("PI-NET実査数入力", "終了")

        if len(not_found_zuban_list) > 0:
            with open("notfound.csv", "a", encoding="UTF-8") as file:
                file.writelines(not_found_zuban_list)
