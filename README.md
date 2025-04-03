## Intro
This is a data action that includes:
1. Reading the dashboard ID
2. Getting each dashboard tile
3. Running the query to get the data for each tile
4. Inserting the data into each Excel tab based on the dashboard tiles.
5. Sending Excel file through SMTP

## Things that need to be accomplished
- Checking each query to see if it has a table calculation or totals. The limit must be lowered to 100,000 rows if either of those are present.
- Comparing the tile filters and the dashboard filters. Tiles can have filters that aren't included on the dashboard.
