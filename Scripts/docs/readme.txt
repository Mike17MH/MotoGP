Workflow to create the final Excel by race using the Analysis.pdf delivered by MotoGP site:

1. pip install pdfplumber xlsxwriter

2. python motogp_parser_generic.py "C:\YourFolder\Analysis.pdf" -o "C:\YourFolder\motogp_race.json"

3. python motogp_excel_builder_generic.py "C:\YourFolder\motogp_race.json" -o "C:\YourFolder\motogp_race.xlsx"

Once we have all the Analysis.pdf parsed and exported to an Excel for the races in the year, 
we can create the Bump Chart Excel using all the JSON generated for each race in the year.
So, please all the JSON in one single folder for a faster and easy handling and execute
the next commands.

1. python motogp_season_bump_chart.py "C:\Users\1021446\Documents\MotoGP\Season_JSON" -o "C:\Users\1021446\Documents\MotoGP\Season_JSON\motogp_season_analysis.xlsx"

Success result message:

Wrote C:\Users\1021446\Documents\MotoGP\Season_JSON\motogp_season_analysis.xlsx: 17 races, 31 season riders

Then now we can create the Probability and Possible Scenarios excel with the next
command:

2. python motogp_probability_and_scenarios.py "C:\Users\1021446\Documents\MotoGP\Season_JSON" -o "C:\Users\1021446\Documents\MotoGP\Season_JSON\motogp_probability_scenario.xlsx" --recent-window 3

Success result message:

Wrote C:\Users\1021446\Documents\MotoGP\Season_JSON\motogp_probability_scenario.xlsx: 17 races, 31 riders
