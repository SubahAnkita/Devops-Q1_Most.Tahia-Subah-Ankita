import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl

def append_suggestions_to_excel(search_query, workbook_path='Keywords.xlsx'):
    
    driver = webdriver.Chrome()  # Make sure 'chromedriver' is in your PATH

    # Open Google
    driver.get('https://www.google.com')

   
    search_box = driver.find_element(By.NAME, 'q')
    search_box.send_keys(search_query)

    # Give some time for auto-suggestions to appear
    time.sleep(2)

    # Find the parent <div> that contains the auto-suggestions
    parent_div = driver.find_element(By.CLASS_NAME, 'erkvQe')

    # Find all the <li> elements inside the parent <div>
    suggestions = parent_div.find_elements(By.TAG_NAME, 'li')

    #determine the longest and shortest strings
    longest_string = ""
    shortest_string = "ThisIsUsedToFindTheMinimumLengthString."
    print("These are the suggested words: \n")
    for suggestion in suggestions:
        text_element = suggestion.find_element(By.TAG_NAME, 'span')
        print(text_element.text)
        string = text_element.text
        if len(string) > len(longest_string):
            longest_string = string
        if len(string) < len(shortest_string):
            shortest_string = string  

    print("-------------------------------------- \n")
    print("Keyword: ", search_query)
    print("Shortest: ", shortest_string)
    print("Longest: ", longest_string)

    # Close the WebDriver
    driver.quit()

    current_day = datetime.datetime.today().strftime('%A')  # e.g., 'Friday'

    # Load the workbook and select the correct sheet based on the day
    workbook = openpyxl.load_workbook(workbook_path)
    sheet = workbook[current_day]  # This will dynamically select the correct sheet

    # Find the next available row in each column
    next_row = sheet.max_row + 1

    #query, longest and shortest suggestions to the next available row in respective columns
    sheet[f'A{next_row}'] = search_query
    sheet[f'B{next_row}'] = longest_string
    sheet[f'C{next_row}'] = shortest_string

    # Save the workbook
    workbook.save(workbook_path)


search_query = input("Enter the search query: ")
append_suggestions_to_excel(search_query)
