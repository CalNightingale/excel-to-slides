import geopandas as gpd
import matplotlib.pyplot as plt
import json
import pandas as pd

def handle_mkt_map(element, provider_data):
    # Read geoJSON
    usa_url = "https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json"
    canada_url = "https://raw.githubusercontent.com/codeforgermany/click_that_hood/main/public/data/canada.geojson"
    usa_gdf = gpd.read_file(usa_url)
    canada_gdf = gpd.read_file(canada_url)
    # Parse down to only desired geography
    contiguous_usa_gdf = usa_gdf[usa_gdf['name'].isin(['Alaska', 'Hawaii', "Puerto Rico"]) == False]
    ontario_gdf = canada_gdf[canada_gdf['name'] == 'Ontario']
    # Shift the geometry of Hawaii to a custom position
    shifted_hawaii = usa_gdf[usa_gdf['name'] == 'Hawaii'].copy()
    shifted_hawaii['geometry'] = shifted_hawaii['geometry'].translate(xoff=45, yoff=5)

    fig, ax = plt.subplots(figsize=(10, 8))
    ax.set_aspect('auto')

    contiguous_usa_gdf.plot(ax=ax, color='lightgray', edgecolor='black')
    shifted_hawaii.plot(ax=ax, color='lightgray', edgecolor='black')
    ontario_gdf.plot(ax=ax, color='lightblue', edgecolor='black')

    plt.axis('off')
    plt.show()



"""
This method is for market presence tables in mystery shopping decks.
These tables have 3 columns: Market, presence?, and quoted?
And have check marks in cells in presence/quoted if we got quotes for that market
To update, we need to 
    (1) remove all existing check marks from presence? and quoted?
    (2) for each market row, check to see if the markets names we pulled from the excel match up
    (3) for each match, add a check mark to both presence? and quoted?
"""
def handle_mkt_presence_table(element, provider_data):
    table = element.Table
    presence_indication_character = "✔"
    for row_index in range(1, len(table.Rows) + 1):
        market_cell = table.Cell(row_index, 1)
        table.Cell(row_index, 2).Shape.TextFrame.TextRange.Text = ""
        table.Cell(row_index, 3).Shape.TextFrame.TextRange.Text = ""
        market_cell_text = market_cell.Shape.TextFrame.TextRange.Text
        for market in provider_data.get("mkt"):
            if market.lower() in market_cell_text.lower():
                table.Cell(row_index, 2).Shape.TextFrame.TextRange.Text = presence_indication_character
                table.Cell(row_index, 3).Shape.TextFrame.TextRange.Text = presence_indication_character
