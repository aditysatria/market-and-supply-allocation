This project implements a Linear Programming (LP) model to optimize the distribution of LPG from SPBEs (LPG Filling Stations) to terminals, minimizing transportation costs and throughput fees. The model is based on demand, capacity, and distance constraints, ensuring that each SPBE is assigned to a single terminal, with the correct quantity shipped.

Features
1. Distance calculation: Uses distance data between SPBEs and terminals which contributes to distribution cost
2. Demand & Capacity: Considers the demand of each SPBE and the capacity of each terminal.
3. Thruput Fee: Takes into account the cost per unit for throughput at each terminal.
4. Terminal cost: Takes into account the cost per unit at each terminal.
5. Supply cost: Takes into account the cost per unit for supply at each terminal.
6. Optimization: The model minimizes the total cost, which includes both transportation and throughput fees.
7. Excel Input/Output: Inputs are provided in an Excel file, and results are output as a list of optimal allocations.

Requirements
1. Python 3.x
2. Libraries: pandas, pulp
3. Excel file format:
Distance to terminal: Distance matrix between SPBEs and terminals.
Demand: The demand from each SPBE.
Alokasi: Terminal capacities.
Thruput Fee: Throughput fee for each terminal.

How to Use
1. Prepare the Excel file with the following sheets:
Distance to terminal: Contains distances between SPBEs and terminals.
Demand: Contains SPBE names and their respective demands.
Alokasi: Contains terminal names, their minimum allocation and their capacities.
Cost: Contains terminal names and their costs.
2. Replace the file path in the script with the path to your Excel file.
3. Run the script:
4. Results will be printed with:
Optimal allocation of LPG from SPBEs to terminals.
Total distance, transport cost, and throughput fees.

Explanation
1. Transportation cost is calculated as distance * transport fee (Rp/ton/km)
2. Thruput fee, supply cost, and terminal cost are the cost per ton transported to the terminal.
