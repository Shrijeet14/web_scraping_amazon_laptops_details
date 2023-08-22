import requests
from bs4 import BeautifulSoup
import csv
import openpyxl

# Define the CSV file to write the data to
csv_filename = "amazon_laptops.csv"

# Create an Excel workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active

# Create a CSV file and write the headers
with open(csv_filename, "w", newline="", encoding="utf-8") as csv_file:
    csv_writer = csv.writer(csv_file)
    csv_writer.writerow(["Title", "URL", "Price", "Ratings"])

    # Define the URL to parse
    for i in range(0,3):
        page_num = i + 1
        url = f"https://www.amazon.in/s?k=laptops&page={page_num}&crid=VNQDLBEWRZZN&qid=1692674794&sprefix=laptop%2Caps%2C227&ref=sr_pg_{page_num}"
        HEADERS = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}

        # Send a GET request to the URL
        response = requests.get(url, headers=HEADERS)
        soup = BeautifulSoup(response.content, "html.parser")

        # Find all the divs with the specified class
        divs = soup.find_all("div", class_="sg-col-inner")

        # Extract data from each div
        for div in divs:
            product_title = div.find("span", class_="a-size-medium a-color-base a-text-normal")
            product_url = div.find("a", class_="a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal")
            product_price = div.find("span", class_="a-price-whole")
            product_ratings = div.find("span", class_="a-icon-alt")

            # Extract the text content or set default values if information is not available
            title = product_title.get_text(strip=True) if product_title else "Information not available"
            url = "https://www.amazon.in" + product_url["href"] if product_url else "Information not available"
            price = product_price.get_text(strip=True) if product_price else "Information not available"
            ratings = product_ratings.get_text(strip=True) if product_ratings else "Information not available"

            # Write the extracted data to the CSV file
            csv_writer.writerow([title, url, price, ratings])
            # Append the extracted data to the sheet
            sheet.append([title, url, price, ratings])

# Save the Excel file
wb.save("amazon_laptops.xlsx")

print("Data has been extracted and saved to amazon_laptops.csv")
