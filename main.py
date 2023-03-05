import requests
from bs4 import BeautifulSoup
URL = 'https://www.imdb.com/chart/top/'
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'IMDb Top 250 Movies'
sheet.append(['Movie Rank','Movie Name','Year Of Release','IMDb Rating'])


try:
    r = requests.get(URL)   
    soup = BeautifulSoup(r.text,'html.parser')
    movies = soup.find('tbody',class_='lister-list').find_all('tr')

    for movie in movies:
        rank = movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        name = movie.find('td',class_='titleColumn').a.text
        year = movie.find('td',class_='titleColumn').span.text.strip('()')
        rating = movie.find('td',class_='ratingColumn imdbRating').strong.text
        sheet.append([int(rank), name, int(year), float(rating)])

except Exception as e:
    print(e)

excel.save('IMDb Top 250 Movies.xlsx')