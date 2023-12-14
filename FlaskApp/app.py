from math import e
from flask import Flask, render_template, request, send_file
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
import os
import time


app = Flask(__name__)

# Global variable to store search results
search_results = None

@app.route("/", methods=["GET", "POST"])
def index():
    global search_results
    message = None  # Initialize a message variable

    if request.method == "POST":
        instructor_name = request.form.get("instructor_name")
        publish_date = request.form.get("PublishDate")

        # Check if the publish date is a valid year
        if publish_date:
            try:
                # Trying to parse the year
                datetime.strptime(publish_date, "%Y")
            except ValueError:
                # If parsing fails, set an error message
                message = "Invalid date format. Please enter a valid year (e.g., 2021)."
                return render_template("search.html", results=None, message=message, excel_file=None)

        search_results = search_google_scholar(instructor_name, publish_date)

        if not search_results:  # Check if the results list is empty
            message = "No results found for the given query."

        return render_template("search.html", results=search_results, message=message, excel_file=(search_results is not None))

    return render_template("search.html", results=None, message=message, excel_file=None)




@app.route('/download_excel/<instructor_name>')
def download_excel(instructor_name):
    global search_results
    if search_results:
        file_path = export_to_excel(search_results, instructor_name)
        return send_file(file_path, as_attachment=True)
    else:
        return "No results to download", 404
    
# def search_google_scholar(instructor_name, publish_date):
#     # Define the URLs for each instructor or use a general query if not listed
# @app.route('/download_excel/<instructor_name>')
def download_excel(instructor_name):
    global search_results
    if search_results:
        file_path = export_to_excel(search_results, instructor_name)
        return send_file(file_path, as_attachment=True)
    else:
        return "No results to download", 404
    
def search_google_scholar(instructor_name, publish_date):
    if instructor_name.lower() == 'tayeb brahimi':
        # This is the URL of Tayeb Brahimi's Google Scholar profile
        url = "https://scholar.google.com/citations?hl=en&user=InKmJHYAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Akila Sarirete':
        url = "https://scholar.google.com/citations?hl=en&user=0Pz559QAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Abdulhamit subasi':
        url = "https://scholar.google.com/citations?hl=en&user=W6FLhskAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Zain Balfagih':
        url = "https://scholar.google.com/citations?hl=en&user=RlxYjNAAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Fidaa Abed':
        url = "https://scholar.google.com/citations?hl=en&user=1Uui0qgAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Sohail Khan':
        url = "https://scholar.google.com/citations?hl=en&user=YygLJOYAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Narjisse Kabbaj':
        url = "https://scholar.google.com/citations?hl=en&user=pEoKxi8AAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Mohammad Nauman':
        url = "https://scholar.google.com/citations?hl=en&user=LgO11BgAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Enfel Barakat':
        url = "https://scholar.google.com/citations?hl=en&user=1HnS9swAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Mohammed Abdulmajid':
        url = "https://scholar.google.com/citations?hl=en&user=ZWd-i_UAAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Amani Ghandour':
        url = "https://scholar.google.com/citations?hl=en&user=OGzqLd0AAAAJ&view_op=list_works&sortby=pubdate"
    elif instructor_name.lower() == 'Nema Salem':
        url = "https://scholar.google.com/citations?hl=en&user=_sDs8l4AAAAJ&view_op=list_works&sortby=pubdate"
    # elif instructor_name.lower() == 'Arbaz Ahmed':
    #     url=""
    # elif instructor_name.lower() == 'Oumaima Geudhami':
    #     url=""
    # elif instructor_name.lower() == 'Mohammed Mousa':
    #     url = ""
    # elif instructor_name.lower() == 'Omar Kitannah':
    #     url=""
    # elif instructor_name.lower() == 'Passant Elkafrawy':
    #     url=""
    # elif instructor_name.lower() == 'Aziza Ibrahim':
    #     url = ""
    else:
        url = f"https://scholar.google.com/scholar?q={instructor_name.replace(' ', '+')}"

    # Set request headers
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    # Perform the request
    time.sleep(5)  # Add a 5-second delay
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        results = []

        for entry in soup.find_all("tr", class_="gsc_a_tr"):
            title = entry.find("a", class_="gsc_a_at").text if entry.find("a", class_="gsc_a_at") else "No Title"
            link = "https://scholar.google.com" + entry.find("a", class_="gsc_a_at")['href'] if entry.find("a", class_="gsc_a_at") else "#"
            date_text = entry.find("span", class_="gsc_a_h").text if entry.find("span", class_="gsc_a_h") else "No Date"

            # Check if the publication date matches the filter
            if date_text != "No Date" and publish_date:
                entry_date = datetime.strptime(date_text, "%Y")
                filter_date = datetime.strptime(publish_date, "%Y")
                if entry_date.year != filter_date.year:
                    continue  # Skip this entry if the year doesn't match

            results.append({"title": title, "date": date_text, "link": link})

        return results
    else:
        print(f"Failed to retrieve data. Status code: {response.status_code}")
        return []

def export_to_excel(results, instructor_name):
    file_path = f"{instructor_name}_research_results.xlsx"

    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Research Results")

    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'Title', bold)
    worksheet.write('D1', 'Date', bold)
    worksheet.write('E1', 'Link', bold)

    for row, data in enumerate(results, start=1):
        worksheet.write(row, 0, data['title'])
        worksheet.write(row, 3, data['date'])
        worksheet.write(row, 4, data['link'])

    workbook.close()
    return file_path

if __name__ == "__main__":
    app.run(debug=True)
