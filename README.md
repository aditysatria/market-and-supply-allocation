This project implements a Linear Programming (LP) model to optimize the distribution of LPG from SPBEs (LPG Filling Stations) to terminals, minimizing transportation costs and throughput fees. The model is based on demand, capacity, and distance constraints, ensuring that each SPBE is assigned to a single terminal, with the correct quantity shipped.

Features
1 Distance calculation: Uses distance data between SPBEs and terminals.
2 Demand & Capacity: Considers the demand of each SPBE and the capacity of each terminal.
3 Thruput Fee: Takes into account the cost per unit for throughput at each terminal.
4 Optimization: The model minimizes the total cost, which includes both transportation and throughput fees.
5 Excel Input/Output: Inputs are provided in an Excel file, and results are output as a list of optimal allocations.

Requirements
1 Python 3.x
2 Libraries: pandas, pulp
3 Excel file format:
Distance to terminal: Distance matrix between SPBEs and terminals.
Demand: The demand from each SPBE.
Alokasi: Terminal capacities.
Thruput Fee: Throughput fee for each terminal.

How to Use
1 Prepare the Excel file with the following sheets:
Distance to terminal: Contains distances between SPBEs and terminals.
Demand: Contains SPBE names and their respective demands.
Alokasi: Contains terminal names and their capacities.
Thruput Fee: Contains terminal names and their throughput fees.
2 Replace the file path in the script with the path to your Excel file.
3 Run the script:
4 Results will be printed with:
Optimal allocation of LPG from SPBEs to terminals.
Total distance, transport cost, and throughput fees.

Explanation
1 Transportation cost is calculated as distance * fuel price * fuel consumption rate.
2 Thruput fee is the cost per ton transported to the terminal.
