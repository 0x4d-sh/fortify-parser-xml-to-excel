import re
import pandas as pd
import argparse
import xlsxwriter
from bs4 import BeautifulSoup

def parse_xml(categories, tab, item_count):
    result = {tab: []}

    for index, category in enumerate(categories):
        index = index + 1

        title = category.find("groupTitle").get_text()
        explanation = category.find("MajorAttributeSummary").find_all("Value")[1].get_text()
        recommendation = category.find("MajorAttributeSummary").find_all("Value")[2].get_text()
        issues = category.find_all("Issue")

        if len(issues) >= item_count:
            result[title], category, risk, description = parse_issues(issues, explanation, recommendation)
            result[tab].append(create_item(index, category, risk, description, explanation, recommendation))
        else:
            tmp, _, _, _ = parse_issues(issues, explanation, recommendation, index, True)
            result[tab] = result[tab] + tmp
            
    return result

def parse_issues(issues, explanation, recommendation, m_index = 0, main = False):
    result = []
    
    category    = issues[0].find("Category").get_text()
    risk        = issues[0].find("Folder").get_text()
    description = issues[0].find("Abstract").get_text()

    for index, issue in enumerate(issues):
        result.append(create_item(
            m_index if main == True else index + 1,
            issue.find("Category").get_text(),
            issue.find("Folder").get_text(),
            issue.find("Abstract").get_text(),
            explanation,
            recommendation,
            issue.find("FilePath").get_text(),
            issue.find("LineStart").get_text(),
            issue.find("Snippet").get_text() if issue.find("Snippet") is not None else "",
            issue.find("TargetFunction").get_text(),
            issue.find("Value").get_text() if issue.find("Value") is not None else ""
        ))

    return result, category, risk, description

def create_item(id, category, risk, description, explanation, recommendation, affected_files = "Please refer to \"{category}\" tab for detailed affected items.", affected_lines = "", snippet = "", highlight = "", analysis = ""):
    custom_list = {}

    custom_list["ID"]               = id
    custom_list["Category"]         = category
    custom_list["Risk"]             = risk
    custom_list["Description"]      = description
    custom_list["Explanation"]      = explanation
    custom_list["Recommendation"]   = recommendation
    custom_list["Affected File(s)"] = affected_files.format(category = category)
    custom_list["Affected Line(s)"] = affected_lines
    custom_list["Snippet"]          = snippet
    custom_list["Highlight"]        = highlight
    custom_list["Analysis"]         = analysis
    
    custom_list["Likelihood"]       = ""
    custom_list["Reference(s)"]     = ""
    custom_list["Consequence"]      = ""
    custom_list["Follow-Up"]        = ""
    return custom_list

def init():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f","--file", help="Set .XML file to import", required=True)
    parser.add_argument("-o","--output", help="Output file name", default="result")
    parser.add_argument("-t","--title", help="Name of the first sheet tab", default="Risk Assessment")
    parser.add_argument("-m","--max-item", help="Max number of data to accept before opening new tab", default=5)
    args = parser.parse_args()

    if args.file.endswith("xml"):
        soup = BeautifulSoup(open(args.file, encoding="utf8"), "xml")
    else:
        print("File type not supported!")
        exit(0)
    
    return args, soup

if __name__ == "__main__":  
    args, soup = init()

    result = parse_xml(soup.find_all("ReportSection")[2].find_all("GroupingSection"), args.title, args.max_item)

    filename = "{file}.xlsx".format(file = args.output)
    writer = pd.ExcelWriter(filename, engine="xlsxwriter")

    for sheet, frame in result.items():
        data = pd.DataFrame(frame)
        data = data[['ID', 'Category', 'Description', 'Affected File(s)', 'Affected Line(s)', 'Snippet', 'Risk', 'Likelihood', 'Consequence', 'Recommendation', 'Reference(s)', 'Follow-Up']]
        data.to_excel(writer, sheet_name = sheet.split(":")[0], index=False)

    writer.save()
    print('[*] Success! Filename: {file}'.format(file = filename))