import requests
from bs4 import BeautifulSoup
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Top Animes'
header_row = ['Rank', 'Title', 'Date Released', 'ratings']
ws.append(header_row)

try:
    source = requests.get('https://www.imdb.com/search/title/?at=0&genres=animation&keywords=anime&num_votes=1000,&sort=user_rating&title_type=tv_series')
    soup = BeautifulSoup(source.text, 'html.parser')

    animes = soup.find('div', class_='lister-list').find_all('div', class_='lister-item mode-advanced')

    for anime in animes:
        path = anime.find('div', class_='lister-item-content').find('h3', class_='lister-item-header')
        rank = path.find('span', class_='lister-item-index unbold text-primary').text   
        # name = anime.find('div', class_='lister-item-content').find('h3', class_='lister-item-header').a.text
        name = path.a.text
        date = path.find('span', class_='lister-item-year text-muted unbold').text.strip('() ')
        
        
        ratings = anime.find('div', class_='lister-item-content').find('div', class_='ratings-bar').find('div', class_='inline-block ratings-imdb-rating').strong.text

        row = [rank, name, date, ratings]
        ws.append(row)
        print(rank, name, date, ratings)
    
    wb.save("C:/Users/Personal Laptop/Desktop/pyProject/data's/top50Anime-iMDB.xlsx") #path file 
        
except Exception as e:
    print(e)
