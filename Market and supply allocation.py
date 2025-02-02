import pandas as pd
from pulp import LpProblem, LpVariable, LpMinimize, lpSum, LpBinary, PULP_CBC_CMD

def optimize_spbe_distribution(file_path):
    # Load data from Excel
    excel_data = pd.ExcelFile(file_path)
    
    # Parse each sheet
    distance_df = excel_data.parse('Distance to terminal')
    demand_df = excel_data.parse('Demand')
    capacity_df = excel_data.parse('Alokasi')
    thruput_fee_df = excel_data.parse('Thruput Fee')
    
    # Extract lists of SPBEs and terminals
    spbe_names = demand_df['PT'].tolist()
    terminal_names = capacity_df['Terminal'].tolist()
    
    # Create dictionaries for demand and capacity
    demand_values = dict(zip(demand_df['PT'], demand_df['Demand']))
    terminal_capacities = dict(zip(capacity_df['Terminal'], capacity_df['Alokasi']))
    thruput_fees = dict(zip(thruput_fee_df['Terminal'], thruput_fee_df['Fee']))
    
    # Create distance matrix
    distances = {}
    for _, row in distance_df.iterrows():
        spbe = row['PT']
        for terminal in terminal_names:
            if terminal in row:
                distances[(spbe, terminal)] = row[terminal]
    
    # Initialize the optimization problem
    model = LpProblem("SPBE_Optimization", LpMinimize)
    
    # Define decision variables
    # X[spbe, terminal] represents the amount shipped
    X = LpVariable.dicts("X", 
                        ((spbe, terminal) for spbe in spbe_names for terminal in terminal_names),
                        lowBound=0,
                        cat='Continuous')
    
    # Y[spbe, terminal] is binary: 1 if terminal serves spbe, 0 otherwise
    Y = LpVariable.dicts("Y",
                        ((spbe, terminal) for spbe in spbe_names for terminal in terminal_names),
                        cat='Binary')
    
    # Objective function: Minimize total transportation cost
    # Objective function: Minimize transportation cost + thruput fee
    model += lpSum(
         Y[spbe, terminal] * distances.get((spbe, terminal), 0) *2*  6800 / 3.1 # Transportation cost, 6800 as fuel price and 3,1 as fuel consumption rate
         + X[spbe, terminal] * 1000 * thruput_fees.get(terminal, 0)     # Thruput fee
         for spbe in spbe_names
         for terminal in terminal_names
    ) 

    
    # Demand constraints
    for spbe in spbe_names:
        model += lpSum(X[spbe, terminal] for terminal in terminal_names) == demand_values[spbe], \
                f"Demand_Constraint_{spbe}"
    
    # Capacity constraints
    for terminal in terminal_names:
        model += lpSum(X[spbe, terminal] for spbe in spbe_names) <= terminal_capacities[terminal], \
                f"Capacity_Constraint_{terminal}"
    
    # Single terminal constraint: Each SPBE must be served by exactly one terminal
    for spbe in spbe_names:
        model += lpSum(Y[spbe, terminal] for terminal in terminal_names) == 1, \
                f"Single_Terminal_{spbe}"
    
    # Linking constraints: X can only be positive if Y is 1
    M = max(demand_values[spbe],terminal_capacities[terminal])
    for spbe in spbe_names:
        for terminal in terminal_names:
            model += X[spbe, terminal] <= M* Y[spbe, terminal], \
                    f"Linking_{spbe}_{terminal}"
    
    # Solve the model
    model.solve(PULP_CBC_CMD(msg=True, timeLimit = 600))  # Maksimal 600 detik, sebelumnya sudah pakai kode & iterasi lama, hasilnya paling bagus muncul di 313 detik
    
    # Return results
    results = []
    total_distance = 0
    total_thruput_fee = 0
    if model.status == 1:  # Optimal solution found
        for spbe in spbe_names:
            for terminal in terminal_names:
                value = X[spbe, terminal].value()
                if value > 0:
                    distance = distances.get((spbe, terminal), 0)
                    fee = thruput_fees.get(terminal, 0)
                    thruput_fee = value * fee
                
                    # Update totals
                    total_distance += distance  # Add 1 for every active route
                    total_thruput_fee += thruput_fee *1000
                
                    results.append({
                        'SPBE': spbe,
                        'Terminal': terminal,
                        'Amount': value,
                        })
    transport_cost = total_distance * 6800 * 2 / 3.1
# Return results
    return results, model.status, total_distance, total_thruput_fee, transport_cost, model.objective.value()

# Usage example:
file_path = "D:\\KP\\Database jarak real.xlsx"
# Usage example:
results, status, total_distance, total_thruput_fee, transport_cost, total_cost = optimize_spbe_distribution(file_path)
total_cost= transport_cost + total_thruput_fee
if status == 1:
    print("\nOptimal Solution Found:")
    for r in results:
        print(f"{r['SPBE']} ,{r['Terminal']}, {r['Amount']:.2f} ton")
    
    # Print totals
    print("\nSummary:")
    print(f"Total Distance: {total_distance}")
    print(f"Transport cost : {transport_cost:.2f}")
    print(f"Total Thruput Fee: {total_thruput_fee:.2f}")
    print(f"Total cost :  {total_cost}")
else:
    print("No optimal solution found. Please check your constraints and data.")