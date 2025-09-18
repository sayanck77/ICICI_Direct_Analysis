import os
import sys
from xml.dom import INVALID_MODIFICATION_ERR
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tabulate import tabulate
import openpyxl
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from datetime import datetime, date, timedelta, time, timezone
from breeze_connect import BreezeConnect
import pytz


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.common.exceptions import NoSuchElementException, WebDriverException, ElementClickInterceptedException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.parse
import time as tm

import smtplib

# MIMEMultipart send emails with both text content and attachments.
from email.mime.multipart import MIMEMultipart
# MIMEText for creating body of the email message.
from email.mime.text import MIMEText
# MIMEApplication attaching application-specific data (like CSV files) to email messages.
from email.mime.application import MIMEApplication

global current_price
global df
global s_token
global inv_df
global pos_df
global ema_rsi_df
global daily_dump_df_list


main_excel_file_path = r"C:\Users\Admin\My_Projects\ICICI_Direct_Analysis\INVESTMENT  - 24-Apr-2025.xlsx"
inv_sheet = "INVESTMENT"
pos_sheet = "POSITIONS"
ema_rsi_sheet = "EMA + RSI"
data_dump_folder = r"C:\Users\Admin\My_Projects\ICICI_Direct_Analysis\DAILY DATA DUMP"
timestamp_file = r"C:\Users\Admin\My_Projects\ICICI_Direct_Analysis\Session_Key\Timestamp.txt"
current_status_folder = r"C:\Users\Admin\My_Projects\ICICI_Direct_Analysis\CURRENT STATUS"
ta_dump = r"C:\Users\Admin\My_Projects\ICICI_Direct_Analysis\TECHNICAL ANALYSIS DUMP"

new_table = {"Stock": [], "Date": [], 'Close': [], "S_EMA1": [], "S_EMA2": [], "C_EMA1" : [], "C_EMA2": [], "S_RSI_P": [], "S_RSI_L_R": [], 'C_RSI_V': [],  'P_Pos': [], 'C_Pos': []}
df = pd.DataFrame(new_table)


def xcel_to_dataframe_creation():
        file = main_excel_file_path
        inv_df = pd.read_excel(file, sheet_name=inv_sheet)
        pos_df = pd.read_excel(file, sheet_name=pos_sheet)
        ema_rsi_df = pd.read_excel(file, sheet_name=ema_rsi_sheet)
        #print(inv_df)
        #print(pos_df)
        #print(ema_rsi_df)
        return inv_df, pos_df, ema_rsi_df



def fetch_stock_list(inv_df):
    stock_list = inv_df["Stock Symbol"].dropna().tolist()
    return stock_list


inv_df, pos_df, ema_rsi_df = xcel_to_dataframe_creation()

fetch_stock_list(inv_df)




