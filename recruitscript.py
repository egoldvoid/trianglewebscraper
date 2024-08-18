# Eli Gold :)
# recruitscript.py
# takes a list of html files from uw directory and converts them to an excel file with a sheet for each department


from bs4 import BeautifulSoup
import pandas as pd

# __________________________________________________ Student Class __________________________________________________

# Student class to store name and number
class Student:
    def __init__(self, name, number):
        self.name = name
        self.number = number

    def __str__(self):
        return self.name + "\t\t" + self.number

    def sex(self, male_names, female_names):
        if self.name.lower().split(' ', 1)[0] in male_names:
            return 0
        elif self.name.lower().split(' ', 1)[0] in female_names:
            return 1
        else:
            return 2

    def getName(self):
        return self.name

    def getNumber(self):
        return self.number

# __________________________________________________ File Processing __________________________________________________

# load male and female names from file
def loadNames(file_path):
    names = set()
    with open(file_path, "r") as file:
        for line in file:
            name = line.strip()
            names.add(name.lower())
    return names


# parse html file and return a list of student objects
def parse_html(file):
    soup = BeautifulSoup(file, 'html.parser')
    persons = []

    for div in soup.find_all('div', class_='rcdescr'):
        header = div.find('h3', class_='scenario-anchor-reference')
        id = header['id'] if header else None

        if id and "students-department-matches" in id:
            for row in soup.find_all('tr', class_='summary-row'):
                name_element = row.find('td', valign='top')
                phone_element = row.find('td', valign='top', nowrap='nowrap')

                if phone_element and '+' in phone_element.get_text(strip=True):
                    name = name_element.get_text(strip=True)
                    number = phone_element.get_text(strip=True)
                    persons.append(Student(name, number))
    return persons


# Convert an html file to three lists of names and numbers
def htmlToLists(filename):
    file = open(filename, "r")
    all_persons = parse_html(file)

    males = [[], []]  # name, numbers
    females = [[], []]
    inconclusive = [[], []]
    lists = [males, females, inconclusive]

    # segregate by sex
    for p in all_persons:
        lists[p.sex(male_names, female_names)][0].append(p.getName())
        lists[p.sex(male_names, female_names)][1].append(p.getNumber())

    # make all lists the same length by adding empty strings
    max_length = max(len(sublist[0]) for sublist in lists)

    lists[0][0].extend([""] * (max_length - len(lists[0][0])))
    lists[0][1].extend([""] * (max_length - len(lists[0][1])))
    lists[1][0].extend([""] * (max_length - len(lists[1][0])))
    lists[1][1].extend([""] * (max_length - len(lists[1][1])))
    lists[2][0].extend([""] * (max_length - len(lists[2][0])))
    lists[2][1].extend([""] * (max_length - len(lists[2][1])))

    return lists

# Convert three lists of names and numbers to a DataFrame
def listsToDataFrame(lists):
    return pd.DataFrame({
        "male names": lists[0][0],
        "male numbers": lists[0][1],
        "empty_1": [""] * len(lists[0][0]),  # first empty column for male
        "empty_2": [""] * len(lists[0][0]),

        "female names": lists[1][0],
        "female numbers": lists[1][1],
        "empty_3": [""] * len(lists[1][0]),  # first empty column for female
        "empty_4": [""] * len(lists[1][0]),

        "inconclusive names": lists[2][0],
        "inconclusive numbers": lists[2][1],
        "empty_5": [""] * len(lists[2][0]), # first empty column for inconclusive
        "empty_6": [""] * len(lists[2][0])
    })


# __________________________________________________ Constants __________________________________________________

# 1000 most common male and female names
male_names = loadNames("./malenames.txt")
female_names = loadNames("./femalenames.txt")

# Sheet names and file names for each department
sheetnames = ['enGrad', 'PreSciences']
filenames = ["./htmls/" + folder + ".html" for folder in sheetnames]


# __________________________________________________ "Main" __________________________________________________
    
# Create excel file with sheets for each department
with pd.ExcelWriter("2425recruitment.xlsx", engine="openpyxl") as writer:
    # add actual sheets
    for sheetname in sheetnames:
        filename = "./htmls/" + sheetname + ".html"
        listsToDataFrame(htmlToLists(filename)).to_excel(
            writer, sheet_name=sheetname, index=False)

