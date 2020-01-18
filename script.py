import xlrd
import xlwt
from requests import request
def cf_ranklist(contest_code):
    guild_handles = {}
    book=xlrd.open_workbook('data.xlsx')
    sheet=book.sheet_by_index(0)
    n=sheet.nrows
    guild_handles = {}
    for i in range(n):
        name = sheet.cell_value(i,2)
        handle = sheet.cell_value(i,5).split('/')[-1]
        guild_handles[handle]=name
    url = "https://codeforces.com/api/contest.ratingChanges?contestId="+str(contest_code)
    page = request('GET',url)
    if not page.ok:
        return []
    data = page.json()
    ranklist = []
    counter = 1
    for row in data['result']:
        username = row['handle']     # works for non team handles else only first handle considered
        if guild_handles.get(username,None):
            ranklist.append((counter,row['rank'],username,guild_handles[username],row['oldRating'],row['newRating']))
            counter += 1
    rank = xlwt.Workbook()
    sheet3 = rank.add_sheet('ranklist')
    
    for i in range(len(ranklist)):
        for j in range(6):
            sheet3.write(i,j,ranklist[i][j])
    rank.save('final.csv')
                   
    #return ranklist 
cf_ranklist(int(input()))
