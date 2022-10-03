from email.charset import SHORTEST
from bs4 import BeautifulSoup
import requests, openpyxl

#create a new excel workbook
excel = openpyxl.Workbook()

#change the sheet name for the active sheet
sheet = excel.active
sheet.title = 'Top 250'

#create the header names to the rows
sheet.append(['Movie Rank', 'Movie Name', 'Release Year', 'IMDB Ranking'])


try:
    #getting the source
    source = requests.get('https://www.imdb.com/chart/top/')

    #cheking the http status to know if the link is correct
    source.raise_for_status()

    #get the html of the link and convert it into html source code
    soup = BeautifulSoup(source.text, 'html.parser')

    #to get the section or table holding all the content you want
    #find_all('tr') will only look for all the tr tag in the tbody tag
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    #loop through the movies body and get all the detail for a particular movie
    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        Releaseyear = movie.find('td', class_="titleColumn").span.text.strip('()')
        Rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
    
        #add the data to the excel sheet
        sheet.append([rank, name, Releaseyear, Rating])

except Exception as e:
    print(e)

#saving the excel work
excel.save('IMDB Top 250 .xlsx')
