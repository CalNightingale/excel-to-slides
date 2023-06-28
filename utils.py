import re
import geopandas as gpd
import matplotlib.pyplot as plt
import json
import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
from PIL import Image

state_abbreviation_to_name = {
    'AL': 'Alabama',
    'AK': 'Alaska',
    'AZ': 'Arizona',
    'AR': 'Arkansas',
    'CA': 'California',
    'CO': 'Colorado',
    'CT': 'Connecticut',
    'DE': 'Delaware',
    'FL': 'Florida',
    'GA': 'Georgia',
    'HI': 'Hawaii',
    'ID': 'Idaho',
    'IL': 'Illinois',
    'IN': 'Indiana',
    'IA': 'Iowa',
    'KS': 'Kansas',
    'KY': 'Kentucky',
    'LA': 'Louisiana',
    'ME': 'Maine',
    'MD': 'Maryland',
    'MA': 'Massachusetts',
    'MI': 'Michigan',
    'MN': 'Minnesota',
    'MS': 'Mississippi',
    'MO': 'Missouri',
    'MT': 'Montana',
    'NE': 'Nebraska',
    'NV': 'Nevada',
    'NH': 'New Hampshire',
    'NJ': 'New Jersey',
    'NM': 'New Mexico',
    'NY': 'New York',
    'NC': 'North Carolina',
    'ND': 'North Dakota',
    'OH': 'Ohio',
    'OK': 'Oklahoma',
    'OR': 'Oregon',
    'PA': 'Pennsylvania',
    'RI': 'Rhode Island',
    'SC': 'South Carolina',
    'SD': 'South Dakota',
    'TN': 'Tennessee',
    'TX': 'Texas',
    'UT': 'Utah',
    'VT': 'Vermont',
    'VA': 'Virginia',
    'WA': 'Washington',
    'WV': 'West Virginia',
    'WI': 'Wisconsin',
    'WY': 'Wyoming',
    'ON': 'Ontario'
}

def handle_mkt_map(slide, element, data):
    present_color = '#104a4a'
    absent_color = '#d4dddf'
    # Parse states from provider data
    states_present = []
    # Multiple states present, loop thru them all
    for state_abbreviation in data.get('state'):
        full_name = state_abbreviation_to_name.get(state_abbreviation)
        if not full_name:
            raise Exception(f"Failed to find full name for state abbreviation '{state_abbreviation}'")
        states_present.append(full_name)

    # Read in geoJSON
    usa_url = "https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json"
    canada_url = "https://raw.githubusercontent.com/codeforgermany/click_that_hood/main/public/data/canada.geojson"
    usa_gdf = gpd.read_file(usa_url)
    canada_gdf = gpd.read_file(canada_url)
    # Parse down to only desired geography
    contiguous_usa_gdf = usa_gdf[usa_gdf['name'].isin(['Alaska', 'Hawaii', "Puerto Rico"]) == False]
    ontario_gdf = canada_gdf[canada_gdf['name'] == 'Ontario']
    hawaii = usa_gdf[usa_gdf['name'] == 'Hawaii']
    # Convert to projection (make it a little curved)
    target_projection = 5070  # Alberts Equal Area Conic Projection
    projected_usa = contiguous_usa_gdf.to_crs(target_projection)
    projected_ontario = ontario_gdf.to_crs(target_projection)
    projected_hawaii = hawaii.to_crs(target_projection)
    # Add presence column for contiguous USA
    projected_usa['presence'] = projected_usa['name'].apply(lambda x: True if x in states_present else False)
    # Shift the geometry of Hawaii to custom position below NM (change offset and rotation to modify position)
    projected_hawaii['geometry'] = projected_hawaii['geometry'].translate(xoff=4600000, yoff=-1400000)
    projected_hawaii['geometry'] = projected_hawaii['geometry'].rotate(35)
    # Plot
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.set_aspect('auto')
    ax.axis('off')
    projected_usa.plot(ax=ax, color=absent_color, edgecolor='white')
    projected_usa[projected_usa['presence']].plot(ax=ax, color=present_color, edgecolor='white')
    projected_hawaii.plot(ax=ax, color=present_color if 'Hawaii' in states_present else absent_color, edgecolor='white')
    projected_ontario.plot(ax=ax, color=present_color if 'Ontario' in states_present else absent_color, edgecolor='white')
    # Save
    plt.savefig('plot.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Replace image
    image_path = os.path.abspath("plot.png")
    new_element = slide.Shapes.AddPicture(FileName=image_path, LinkToFile=False,
                                          SaveWithDocument=True, Left=element.left,
                                          Top=element.top, Width=element.width, Height=element.height)
    new_element.name = element.name
    element.Delete()


"""
This method is for market presence tables in mystery shopping decks.
These tables have 3 columns: Market, presence?, and quoted?
And have check marks in cells in presence/quoted if we got quotes for that market
To update, we need to 
    (1) remove all existing check marks from presence? and quoted?
    (2) for each market row, check to see if the markets names we pulled from the excel match up
    (3) for each match, add a check mark to both presence? and quoted?
"""
def handle_mkt_presence_table(slide, element, data):
    table = element.Table
    presence_indication_character = "✔"
    for row_index in range(1, len(table.Rows) + 1):
        market_cell = table.Cell(row_index, 1)
        table.Cell(row_index, 2).Shape.TextFrame.TextRange.Text = ""
        table.Cell(row_index, 3).Shape.TextFrame.TextRange.Text = ""
        market_cell_text = market_cell.Shape.TextFrame.TextRange.Text
        for market in data.get("mkt"):
            if market.lower() in market_cell_text.lower():
                table.Cell(row_index, 2).Shape.TextFrame.TextRange.Text = presence_indication_character
                table.Cell(row_index, 3).Shape.TextFrame.TextRange.Text = presence_indication_character

# Helper method to get aspect ratio of image
def get_aspect_ratio(image_path):
    with Image.open(image_path) as image:
        width, height = image.size
        aspect_ratio = width / height
        return aspect_ratio

"""
This method is for fetching provider logos from the internet automatically.
It does this by getting the link to the first image returned by searching for 
"{provider name} logo" on google images with the transparent background tool enabled.
It then downloads this image and updates the logo element accordingly
"""
def handle_logo(slide, element, data):
    # Generate http request
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
    }
    company_name = data.get('provider')[0]
    params = {    
        "q": f"{company_name}+logo",    # search query
        "tbm": "isch",                  # image results
        "tbs": "ic:trans",              # transparent results only
        "hl": "en",                     # language of the search
        "gl": "us"                      # country where search comes from
    }
    # Send request
    html = requests.get("https://google.com/search", params=params, headers=headers, timeout=30)

    # Parse result
    soup = BeautifulSoup(html.text, "lxml")
    all_script_tags = soup.select("script")
    # https://regex101.com/r/RPIbXK/1
    matched_images_data = "".join(re.findall(r"AF_initDataCallback\(([^<]+)\);", str(all_script_tags)))
    matched_images_data_fix = json.dumps(matched_images_data)
    matched_images_data_json = json.loads(matched_images_data_fix)

    # https://regex101.com/r/NRKEmV/1
    matched_google_image_data = re.findall(r'\"b-GRID_STATE0\"(.*)sideChannel:\s?{}}', matched_images_data_json)
                
    # Remove previously matched thumbnails for easier full resolution image matches
    removed_matched_google_images_thumbnails = re.sub(
            r'\[\"(https\:\/\/encrypted-tbn0\.gstatic\.com\/images\?.*?)\",\d+,\d+\]', "", str(matched_google_image_data))
        
    # https://regex101.com/r/fXjfb1/4
    # https://stackoverflow.com/a/19821774/15164646
    matched_google_full_resolution_images = re.findall(r"(?:'|,),\[\"(https:|http.*?)\",\d+,\d+\]", removed_matched_google_images_thumbnails)

    full_res_images = [
            bytes(bytes(img, "ascii").decode("unicode-escape"), "ascii").decode("unicode-escape") for img in matched_google_full_resolution_images
    ]
    # Choose the first image returned
    image_data = requests.get(full_res_images[0]).content
    with open('logo.png', 'wb') as writer:
        writer.write(image_data)
    # Replace image
    image_path = os.path.abspath("logo.png")
    aspect_ratio = get_aspect_ratio("logo.png")
    # Image should be the same width as the template logo but should maintain aspect ratio
    new_element = slide.Shapes.AddPicture(FileName=image_path, LinkToFile=False,
                                          SaveWithDocument=True, Left=element.left,
                                          Top=element.top, Width=element.width,
                                          Height=element.width / aspect_ratio)
    new_element.name = element.name
    element.Delete()
