import os
import signal
from tkinter import filedialog, messagebox

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import jissa_file


def login():
    driver.get(os.getenv("url"))

    user_name_box = driver.find_element(By.NAME, "username")
    user_name_box.send_keys(os.getenv("user_name"))

    password_box = driver.find_element(By.NAME, "password")
    password_box.send_keys(os.getenv("password"))

    loggin_button = driver.find_element(By.XPATH, "/html/body/form/input[2]")
    loggin_button.click()


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


def click_screen_update_button():
    """
    画面右上の検索/更新ボタンを実行する
    """

    driver.execute_script("javascript:clickButton(''); return top.frSubMenu.clickExecute();")


def input_jissa_su() -> bool:
    """
    pinetの対象図番を実査登録確認リストから取得し、実査数をpinetに入力する。
    図番がなくなったら終了し、Falseを返す
    :return: bool
    """
    zuban = ""
    for row in range(2, 16):
        zuban_td = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[2]")
        zuban = zuban_td.text.strip()

        if zuban == "":
            break

        jissa_su: float = jissa.get_jissa_su(zuban)
        if jissa_su is not None:
            zaikosu = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[8]/input")
            zaikosu.clear()
            zaikosu.send_keys(jissa_su)

            checkbox = driver.find_element(By.XPATH, f"/html/body/form/table[7]/tbody/tr[{row}]/td[1]/input[2]")
            driver.execute_script("arguments[0].click();", checkbox)
    return zuban != ""


def click_next_page():
    driver.execute_script("javascript:return top.frSubMenu.clickNextPage();")


if __name__ == "__main__":
    load_dotenv()

    file_type = [("Excel File", "*.xlsx")]
    base_dir = os.getenv("base_dir")
    file_path = filedialog.askopenfilename(filetypes=file_type, initialdir=base_dir)

    if file_path != "":
        jissa = jissa_file.JissaFile(file_path)

        driver = webdriver.Edge()
        driver.implicitly_wait(10)

        login()
        change_to_input_frame()
        click_screen_update_button()
        while input_jissa_su():
            click_next_page()

        os.kill(driver.service.process.pid, signal.SIGTERM)
        messagebox.showinfo("PI-NET実査数入力", "完了")
