import requests
from bs4 import BeautifulSoup
import pandas as pd

import openpyxl

import re


def writeDataInExcel(cyclists_data, path_to_write):
    df = pd.DataFrame(data = cyclists_data)
    #df.to_excel("./myData.xlsx", index=False)
    df.to_excel(path_to_write, index=False)
    print("Data was written to excel")

def getCyclistsLinksMap(file_path, sheet_name, column_letter):
    
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the sheet by name
        sheet = workbook[sheet_name]

        # Get the column by letter
        column = sheet[column_letter]

        # Initialize a list to store hyperlinks
        hyperlinks_map = {}

        # Iterate through each cell in the column
        for cell in column:
            print(cell.value)
            hyperlink_element = cell.hyperlink.target if cell.hyperlink else None
            hyperlinks_map[cell.value] = hyperlink_element

        return hyperlinks_map

    except Exception as e:
        print(f"Error: {e}")
        return None
    


def scrape_cyclist_data(urls):
    # Send a GET request to the URL

    cyclists_infos = []
    for key,value in urls.items(): 

        response = requests.get(value)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Parse the HTML content of the page
            soup = BeautifulSoup(response.text, 'html.parser')

            #main section (name)
            main_section = soup.find("div", class_="main")
            h1_tag_name = main_section.find("h1")
            first_last_name = h1_tag_name.text
            print(first_last_name)

            #Age
            rdr_info_content  = soup.find("div", class_="rdr-info-cont")
            print(rdr_info_content)
            print("*****************************************************************************")
            rdr_text_content = rdr_info_content.get_text(" ", strip=True)
            age_match = re.search(r'\((\d+)\)', rdr_text_content)
            age = age_match.group(1)

            #Nationality
            nationality_tag = rdr_info_content.find('a', class_='black')
            nationality = nationality_tag.text

            #height and weight
            weight_element = soup.find('b', string='Weight:')
            height_element = soup.find('b', string='Height:')
 
            #points
            points_elements = soup.find_all("div", class_='pnt')
            
            one_day_races = points_elements[0].text
            gc_points = points_elements[1].text
            time_trial_points = points_elements[2].text
            sprint_points = points_elements[3].text
            climber_points = points_elements[4].text

            #ranking
            ranking_points = soup.find_all("div", class_='rnk')

            uci_ranking = ""
            pcs_ranking = ""
            all_time_ranking = ""
            
            uci_index = 0
            pcs_index = 1
            all_time_index = 2

            len_ranking_points = len(ranking_points)

            if uci_index < len_ranking_points :
                uci_ranking = ranking_points[0].text
            if pcs_index < len_ranking_points:
                pcs_ranking = ranking_points[1].text
            if all_time_index < len_ranking_points: 
                all_time_ranking = ranking_points[2].text
            
            #print(ranking_points)
            
            weight_element_sibling = ""
            height_element_sibling = ""

            if weight_element != None and height_element != None: 
                weight_element_sibling = weight_element.next_sibling.strip()
                height_element_sibling = height_element.next_sibling.strip()
        
            #return {'Height': height, 'Weight': weight}

            #return {'rider_infos': rider_infos} 
            
            #Transform to numbers
            map_data = {'Name': key, 
                        'Nationality': nationality,
                        'Weight': weight_element_sibling, 
                        'Height': height_element_sibling, 
                        'Age': age,
                        'One day races': one_day_races,
                        'GC': gc_points,
                        'Time trial': time_trial_points,
                        'Sprint': sprint_points,
                        'Climber': climber_points,
                        "UCI Ranking": uci_ranking,
                        "PCS Ranking": pcs_ranking,
                        "All time Ranking": all_time_ranking,
                        } 

            #add the map to the global list of cyclists
            cyclists_infos.append(map_data)

        else:
            print(f"Failed to retrieve data. Status Code: {response.status_code}")
            return None
    
    return cyclists_infos


excel_file_path = "./Data2023.xlsx"
sheet_name = "Sheet5"
column_letter = "A"


hyperlinks_list = getCyclistsLinksMap(excel_file_path, sheet_name, column_letter)
result = scrape_cyclist_data(hyperlinks_list)
writeDataInExcel(result, "./DayZ_Data.xlsx")
