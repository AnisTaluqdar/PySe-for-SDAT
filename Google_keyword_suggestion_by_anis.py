#Author: Muhammd Anisur Rahman

# Importing necessary module
import time
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# Making a function for keyword pass
def search_keyword(keyword):
    driver = webdriver.Firefox()
    driver.get("https://www.google.com")
    search_box = driver.find_element(By.NAME, "q")
    search_box.send_keys(keyword)

    # Sleeping for visualize or getting data in search box
    time.sleep(3)

    # Finding the suggestion box data in google search by ID
    suggestion_elements = driver.find_elements(By.ID,"Alh6id")
    for element in suggestion_elements:
        suggestions = element.text.splitlines() # Splitting the line to as a list

    # For no suggestions result taking care the error
    if not suggestions:
        longest_option = "No suggestions found"
        shortest_option = "No suggestions found"
    else:
        longest_option = max(suggestions, key=len) # Find out the maximum value
        shortest_option = min(suggestions, key=len) # Find out the minimum value

    driver.quit()
    return longest_option, shortest_option

# Defining the today's date
today = datetime.datetime.now().strftime("%A")
excel_file = "C:\\Users\\anist\\Downloads\\Documents\\file.xlsx"
df = pd.read_excel(excel_file, sheet_name=today)
keywords = df["Unnamed: 2"].tolist()

for i, keyword in enumerate(keywords):
    longest_option, shortest_option = search_keyword(keyword) # Call search_keyword function
    if i < 1: # passing first data because of the empty value given in excel data
        pass
    else:
        df.loc[i, "Unnamed: 3"] = longest_option # Storing longest value locationwise
        df.loc[i, "Unnamed: 4"] = shortest_option # Storing shortest value locationwise

# Writing the excel file using pandas function
with pd.ExcelWriter(excel_file,  engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name=today, index=False, header=False, startrow=1)
