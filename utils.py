"""
This method is for market presence tables in mystery shopping decks.
These tables have 3 columns: Market, presence?, and quoted?
And have check marks in cells in presence/quoted if we got quotes for that market
To update, we need to 
    (1) remove all existing check marks from presence? and quoted?
    (2) for each market row, check to see if the markets names we pulled from the excel match up
    (3) for each match, add a check mark to both presence? and quoted?
"""
def handle_mkt_presence_table(table, provider_data):
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
