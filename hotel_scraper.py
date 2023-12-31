"""
Script: hotel_scraper.py
Author: Brandon Short
Date: December 30, 2023

Description:
This script uses the Playwright library to scrape hotel information from Booking.com.
It prompts the user for destination, check-in and check-out dates, number of adults, children, rooms, and desired result pages.
Integrated Pandas for data manipulation and created a Pandas DataFrame to organize and structure the scraped data.
The collected data is then saved to an Excel file named 'hotels_list.xlsx' using Pandas.
Used XlsxWriter to save the collected data in an Excel spreadsheet for easy sharing and analysis.

Dependencies:
- Playwright (install using: pip/pip3 install playwright)
- Pandas (install using: pip/pip3 install pandas)

Usage:
1. Run the script.
2. Enter the required information when prompted.
3. The script will scrape hotel data from Booking.com and save it to 'hotels_list.xlsx'.

Note:
Make sure to install the required dependencies before running the script.
"""

from playwright.sync_api import sync_playwright
import pandas as pd
import datetime

## Input Validation ##
# Date
def validDate(date):
    try:
        datetime.date.fromisoformat(date)
        return True
    except ValueError:
        print("Incorrect data format, should be YYYY-MM-DD")
        return False

# Number of adults, children, and rooms
def validAdults(numAdults):
    if numAdults > 30 or numAdults < 1:
        print("Number entered is out of range.")
        return False
    else:
        return True

def validChildren(numChildren):
    if numChildren > 10 or numChildren < 0:
        print("Number entered is out of range.")
        return False
    else:
        return True

def validRooms(numRooms):
    if numRooms > 30 or numRooms < 1:
        print("Number entered is out of range.")
        return False
    else:
        return True

# Number of pages
def validPages(numPages):
        if numPages > 8 or numPages < 1:
            print("Number entered is out of range.")
            return False
        else:
            return True

# Go to the next page
def nextPage(checkin_date, checkout_date, location, num_adults, num_rooms, num_children, page, existingPage):
    offset =+ int(existingPage) * 25
    page_url = f"https://www.booking.com/searchresults.en-us.html?checkin={checkin_date}&checkout={checkout_date}&selected_currency=USD&ss={location}&ssne={location}&ssne_untouched={location}&lang=en-us&sb=1&src_elem=sb&src=searchresults&dest_type=city&group_adults={num_adults}&no_rooms={num_rooms}&group_children={num_children}&sb_travel_purpose=leisure&offset={offset}"
    page.goto(page_url, timeout=60000)
    hotels = page.locator('//div[@data-testid="property-card"]').all()
    
    return hotels

# Add hotels to the list, from the page
def addHotels(hotels, hotelsList):
    for hotel in hotels:
        hotelDict = {}
        hotelDict['Hotel Name'] = hotel.locator('//div[@data-testid="title"]').inner_text()
        hotelDict['Price'] = hotel.locator('//span[@data-testid="price-and-discounted-price"]').inner_text()
        hotelDict['Score'] = hotel.locator('//div[@data-testid="review-score"]/div[1]').inner_text()
        hotelDict['Avg review'] = hotel.locator('//div[@data-testid="review-score"]/div[2]/div[1]').inner_text()
        hotelDict['Reviews count'] = hotel.locator('//div[@data-testid="review-score"]/div[2]/div[2]').inner_text().split()[0]

        hotelsList.append(hotelDict)
    
    return hotelsList

def main():
    with sync_playwright() as p:
        ### Get the Destination and elements of the trip ###
        # Get the location and check if it's valid
        location = input("Enter your Destination: ")

        # Get the dates and check if they are valid
        checkin_date = input("Enter Check-In Date (YYYY-MM-DD): ")
        while True:
            if validDate(checkin_date) == True:
                break
            else: 
                checkin_date = input("Enter Check-In Date (YYYY-MM-DD): ")

        checkout_date = input("Enter Check-Out Date (YYYY-MM-DD): ")
        while True:
            if validDate(checkout_date) == True:
                break
            else: 
                checkout_date = input("Enter Check-Out Date (YYYY-MM-DD): ")

        # Get the number of adults and check if it's valid
        num_adults = int(input("Enter the # of adults (Max. 30): "))
        while True:
            try:
                if validAdults(num_adults) == True:
                    break
                else:
                    num_adults = int(input("Enter the # of adults (Max. 30): "))
            except ValueError:
                print("Please enter a valid number.")
                num_adults = int(input("Enter the # of adults (Max. 30): "))

        # Get the number of children and check if it's valid
        num_children = int(input("Enter the # of children (Max. 10): "))
        while True:
            try:
                if validChildren(num_children) == True:
                    break
                else:
                    num_children = int(input("Enter the # of children (Max. 10): "))
            except ValueError:
                print("Please enter a valid number.")
                num_children = int(input("Enter the # of children (Max. 10): "))

        # Get the number of rooms and check if it's valid
        num_rooms = int(input("Enter the # of rooms (Max. 30): "))
        while True:
            try:
                if validRooms(num_rooms) == True:
                    break
                else:
                    num_rooms = int(input("Enter the # of rooms (Max. 30): "))
            except ValueError:
                print("Please enter a valid number.")
                num_rooms = int(input("Enter the # of rooms (Max. 30): "))

        # Get the number of pages the user wants to see and check if it's vaid
        num_pages = int(input("How many pages of results would you like to see? (Max. 8) "))
        while True:
            try:
                if validPages(num_pages) == True:
                    break
                else:
                    num_pages = int(input("How many pages of results would you like to see? (Max. 8) "))
            except ValueError:
                print("Please enter a valid number.")
                num_pages = int(input("How many pages of results would you like to see? (Max. 8) "))

        # Input the elements into a URL
        page_url = f"https://www.booking.com/searchresults.en-us.html?checkin={checkin_date}&checkout={checkout_date}&selected_currency=USD&ss={location}&ssne={location}&ssne_untouched={location}&lang=en-us&sb=1&src_elem=sb&src=searchresults&dest_type=city&group_adults={num_adults}&no_rooms={num_rooms}&group_children={num_children}&sb_travel_purpose=leisure"

        # Open Chromium and go to Booking.com
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(page_url, timeout=60000)

        # Locate the HTML element for the hotel card
        hotels = page.locator('//div[@data-testid="property-card"]').all()

        # Create an empty list for the hotels
        hotelsList = []
        # For each hotel on the page add it and it's elements to the hotel list
        hotelsList = addHotels(hotels, hotelsList)

        # Loop over the # of pages entered by the user, then add to the list 
        for i in range(1, num_pages):
            hotels = nextPage(checkin_date, checkout_date, location, num_adults, num_rooms, num_children, page, i) 
            hotelsList = addHotels(hotels, hotelsList)

        # Print the amount of hotels at the location
        print(f'Your {len(hotelsList)} hotels in {location} are ready to be viewed in Excel.')

        # Create the Pandas data frame
        df = pd.DataFrame(hotelsList)

        # Create the Pandas Excel writer using the XlsxWriter as te engine
        writer = pd.ExcelWriter('hotels_list.xlsx', engine='xlsxwriter')

        # Convert the DataFrame into an XlsxWriter Excel Object
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)

        # Get the XlsxWriter objects from the dataframe writer object
        worksheet = writer.sheets['Sheet1']

        # Set column width
        worksheet.autofit()
        # Close the writer (and save) and browser
        writer.close()
        browser.close()

if __name__ == "__main__":
    main()