import openpyxl


categories = [5,18,20,23,24,26,27,28,35]

def key(e):
    return e[0]

def format(value):
    if "£" in value:
        value.replace("£", "")
    elif "%" in value:
        value.replace("%", "")

    return value



def lowerQuartile(set):
    Q3 = len(set)*0.75
    Q3 = round(Q3)
    return int(Q3)

class RetailCentre:
    def __init__(self):
        self.DataFields = []
        self.points = 0

def loadData():
    workbook = openpyxl.load_workbook(filename="Pure Data.xlsx", read_only=True)
    sheet = workbook["RF Data"]
    retailCentres = []
    for row in sheet:
        new = RetailCentre()
        for cell in row:
            if cell.value != None:
                poop = cell.value
                new.DataFields.append(poop)
        if len(new.DataFields) > 0:
            retailCentres.append(new)
    return retailCentres

def RankBy(catInd, data, weighting=1):

    values = []
    for i in data:
        values.append(i.DataFields[catInd])
    values.sort()
    LQ = lowerQuartile(values)
    d2 = data
    for d in d2:
        d.points += (d.DataFields[catInd]/values[LQ])*100*weighting
    return d2

def displayPoints(data, index = -1):
    if index == -1:
        for d in data:
            print(d.points)
    else:
        print(data[index].points)
def competitiors(data):
    for d in data:
        d.points = d.points/(1+d.DataFields[45]+d.DataFields[43])
    return data
data = loadData()

def PointsRank(data):
    points = []
    for i in data:
        points.append([i.points, i.DataFields[0], i.DataFields[1]])
    points.sort(key=key, reverse=True)
    return points




for x in categories:
    data=RankBy(x, data)

displayPoints(data, 0)

data = competitiors(data)

final = PointsRank(data)
print(final)
