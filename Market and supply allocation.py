import pandas as pd
from pulp import LpProblem, LpVariable, LpMinimize, lpSum, LpBinary, PULP_CBC_CMD

# Usage example:
file_path = "D:\\KP\\Database jarak real.xlsx"
excel_data = pd.ExcelFile(file_path)
    
    # Parse each sheet
distance_df = excel_data.parse('Distance data pertamina')
demand_df = excel_data.parse('Demand Average')
capacity_df = excel_data.parse('Alokasi')
cost_df = excel_data.parse('Cost')

    # Extract lists of SPBEs and terminals
spbe_names = demand_df['PT'].tolist()
terminal_names = capacity_df['Terminal'].tolist()
    
    # Create dictionaries for demand and capacity
demand_values = dict(zip(demand_df['PT'], demand_df['Demand']))
terminal_allocation = dict(zip(capacity_df['Terminal'], capacity_df['Alokasi']))
terminal_capacities = dict(zip(capacity_df['Terminal'], capacity_df['Kapasitas']))
terminal_costs = dict(zip(cost_df['Terminal'], cost_df['Terminal Cost']))
thruput_fees = dict(zip(cost_df['Terminal'], cost_df['Thruput Fee']))
supply_costs = dict(zip(cost_df['Terminal'], cost_df['Supply Cost']))
    
# Create distance matrix
distances = {}
for _, row in distance_df.iterrows():
    spbe = row['PT']
    for terminal in terminal_names:
        if terminal in row:
            distances[(spbe, terminal)] = row[terminal]
    
single_sourcing_terminals = ["TERMINAL LPG REMBANG", "DEPOT LPG BALONGAN", "DEPOT LPG CILACAP"]
multi_sourcing_terminals = ["OPSICO TERMINAL LPG SEMARANG", "TERMINAL LPG PEL SEMARANG"]

obligatory_terminals = ["DEPOT LPG BALONGAN", "DEPOT LPG CILACAP"]
active_terminal = 5  # Ganti dengan jumlah terminal aktif yang diinginkan

single_sourcing_terminals = [terminal for terminal in terminal_names if terminal in single_sourcing_terminals]
multi_sourcing_terminals = [terminal for terminal in terminal_names if terminal in multi_sourcing_terminals]

def optimize_spbe_distribution(file_path):
    
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

    
    # Objective function: Minimize total cost
    model += lpSum(
         X[spbe, terminal] * 
         ((distances.get((spbe, terminal), 0) *1454) # Transportation cost
        + terminal_costs.get(terminal,0) * 1000 #terminal cost
        + thruput_fees.get(terminal, 0) * 1000 #thruput fee
        + supply_costs.get(terminal,0) * 1000)     # Thruput fee
         for spbe in spbe_names
         for terminal in terminal_names
    ) 
    
    # Memastikan setiap obligatory terminal dipilih oleh setidaknya satu SPBE
    for terminal in obligatory_terminals:
        model += lpSum(Y[spbe, terminal] for spbe in spbe_names) >= 1, f"Force_{terminal}_Chosen"
        model += lpSum(X[spbe, terminal] for spbe in spbe_names) >= terminal_allocation[terminal], f"Capacity_{terminal}_Check"

    # SPBE hanya bisa mendapatkan pasokan dari satu sumber single-sourcing atau dua multi-sourcing
    for spbe in spbe_names:
        # Demand constraints
        model += lpSum(X[spbe, terminal] for terminal in terminal_names) == demand_values[spbe], \
            f"Demand_Constraint_{spbe}"

    for spbe in spbe_names:    
        # 1. Exactly ONE or TWO terminals
        model += lpSum(Y[spbe, terminal] for terminal in terminal_names) <= 2, f"Max_Two_Terminals_{spbe}"
        model += lpSum(Y[spbe, terminal] for terminal in terminal_names) >= 1, f"Min_One_Terminals_{spbe}"

        # 2. Indicator variables for single and multi sourcing
        single_sourcing_chosen = LpVariable(f"single_sourcing_chosen_{spbe}", cat='Binary')
        multi_sourcing_chosen = LpVariable(f"multi_sourcing_chosen_{spbe}", cat='Binary')

        # 3. Link indicator variables to Y
        model += lpSum(Y[spbe, terminal] for terminal in single_sourcing_terminals) <= len(single_sourcing_terminals) * single_sourcing_chosen, f"Link_Single_{spbe}"
        model += lpSum(Y[spbe, terminal] for terminal in multi_sourcing_terminals) <= len(multi_sourcing_terminals) * multi_sourcing_chosen, f"Link_Multi_{spbe}"

        # 4. Prevent mixed sourcing
        model += single_sourcing_chosen + multi_sourcing_chosen <= 1, f"No_Mix_{spbe}"

        # 5. Limit single sourcing terminals if chosen
        model += lpSum(Y[spbe, terminal] for terminal in single_sourcing_terminals) <= 1 * single_sourcing_chosen, f"Limit_Single_{spbe}"

        # 6. Limit multi sourcing terminals if chosen
        model += lpSum(Y[spbe, terminal] for terminal in multi_sourcing_terminals) <= 2 * multi_sourcing_chosen, f"Limit_Multi_{spbe}"    
        
    is_terminal_active = {}  # Inisialisasi dictionary di sini

    # Capacity constraints
    for terminal in terminal_names:
        # Variabel biner untuk menunjukkan apakah terminal aktif atau tidak (sekarang di dalam loop)
        is_terminal_active[terminal] = LpVariable(f"is_terminal_active_{terminal}", cat='Binary')

        # Constraint yang menghubungkan variabel biner dengan Y (apakah ada SPBE yang memilih terminal)
        model += lpSum(Y[spbe, terminal] for spbe in spbe_names) <= len(spbe_names) * is_terminal_active[terminal], f"Link_Active_{terminal}"

        # Constraint alokasi minimum yang dimodifikasi
        model += lpSum(X[spbe, terminal] for spbe in spbe_names) >= terminal_allocation[terminal] * is_terminal_active[terminal], f"Min_Allocation_{terminal}"

        model += lpSum(X[spbe, terminal] for spbe in spbe_names) <= terminal_capacities[terminal], f"Capacity_Constraint_{terminal}"

    # Batasan jumlah terminal aktif (tepat)
    model += lpSum(is_terminal_active[terminal] for terminal in terminal_names) >= active_terminal, "Tepat_Jumlah_Terminal_Aktif"

# ... (sisa kode) ...
    # Linking constraints: X can only be positive if Y is 1M = max(demand_values[spbe] for spbe in spbe_names)  # Pastikan M cukup besar
    M = max(demand_values.values())  # Atau nilai besar lainnya
    for spbe in spbe_names:
        for terminal in terminal_names:
            model += X[spbe, terminal] <= M * Y[spbe, terminal]


    # Solve the model
    model.solve(PULP_CBC_CMD(msg=True))

    # Return results
    results = []
    total_distance = 0
    total_terminal_cost = 0
    total_thruput_fee = 0
    total_supply_cost = 0
    total_transport_cost = 0
    if model.status == 1:  # Optimal solution found
        for spbe in spbe_names:
            for terminal in terminal_names:
                value = X[spbe, terminal].value()
                if value > 0:
                    distance = distances.get((spbe, terminal), 0)
                    transport_cost = distances.get((spbe, terminal), 0) * value * 1454
                    terminal_cost = terminal_costs.get(terminal,0) * value * 1000
                    thruput_fee = thruput_fees.get(terminal, 0) * value * 1000
                    supply_cost = supply_costs.get(terminal,0) * value * 1000
                
                    # Update totals
                    total_distance += distance  # Add 1 for every active route
                    total_terminal_cost += terminal_cost
                    total_thruput_fee += thruput_fee
                    total_supply_cost += supply_cost
                    total_transport_cost += transport_cost
                
                    results.append({
                        'SPBE': spbe,
                        'Terminal': terminal,
                        'Amount': value,
                        })
            
# Return results
    return results, model.status, total_distance, total_terminal_cost, total_thruput_fee, total_supply_cost, total_transport_cost, model.objective.value()

# Usage example:
results, status, total_distance, total_terminal_cost, total_thruput_fee, total_supply_cost, total_transport_cost, total_cost = optimize_spbe_distribution(file_path)
total_cost= total_terminal_cost + total_thruput_fee + total_supply_cost + total_transport_cost

total_dot_opsico = sum(r['Amount'] for r in results if r['Terminal'] == "OPSICO TERMINAL LPG SEMARANG") 
total_dot_pel = sum(r['Amount'] for r in results if r['Terminal'] == "TERMINAL LPG PEL SEMARANG")
total_dot_rembang = sum(r['Amount'] for r in results if r['Terminal'] == "TERMINAL LPG REMBANG") 
total_dot_balongan = sum(r['Amount'] for r in results if r['Terminal'] == "DEPOT LPG BALONGAN")
total_dot_cilacap = sum(r['Amount'] for r in results if r['Terminal'] == "DEPOT LPG CILACAP")

total_thruput_opsico = sum(r['Amount'] * thruput_fees[r['Terminal']] * 1000 
                           for r in results if r['Terminal'] == "OPSICO TERMINAL LPG SEMARANG")
total_thruput_pel = sum(r['Amount'] * thruput_fees[r['Terminal']] * 1000 
                        for r in results if r['Terminal'] == "TERMINAL LPG PEL SEMARANG")
total_thruput_rembang = sum(r['Amount'] * thruput_fees[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "TERMINAL LPG REMBANG")

total_terminal_cost_balongan = sum(r['Amount'] * terminal_costs[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "DEPOT LPG BALONGAN")
total_terminal_cost_cilacap = sum(r['Amount'] * terminal_costs[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "DEPOT LPG CILACAP")

total_supply_cost_opsico = sum(r['Amount'] * supply_costs[r['Terminal']] * 1000 
                           for r in results if r['Terminal'] == "OPSICO TERMINAL LPG SEMARANG")
total_supply_cost_pel = sum(r['Amount'] * supply_costs[r['Terminal']] * 1000 
                        for r in results if r['Terminal'] == "TERMINAL LPG PEL SEMARANG")
total_supply_cost_rembang = sum(r['Amount'] * supply_costs[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "TERMINAL LPG REMBANG")
total_supply_cost_balongan = sum(r['Amount'] * supply_costs[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "DEPOT LPG BALONGAN")
total_supply_cost_cilacap = sum(r['Amount'] * supply_costs[r['Terminal']] * 1000 
                            for r in results if r['Terminal'] == "DEPOT LPG CILACAP")

if status == 1:
    print("\nOptimal Solution Found:")
    for r in results:
        print(f"{r['SPBE']} ,{r['Terminal']}, {r['Amount']:.2f}")
    
    print(f"\nTotal DOT untuk OPSICO: {total_dot_opsico:.2f}")
    print(f"Total DOT untuk PEL: {total_dot_pel:.2f}")
    print(f"Total DOT untuk REMBANG: {total_dot_rembang:.2f}")
    print(f"Total DOT untuk BALONGAN: {total_dot_balongan:.2f}")
    print(f"Total DOT untuk CILACAP: {total_dot_cilacap:.2f}")

    print(f"\nTotal Thruput Fee untuk OPSICO: {total_thruput_opsico:.2f}")
    print(f"Total Thruput Fee untuk PEL: {total_thruput_pel:.2f}")
    print(f"Total Thruput Fee untuk REMBANG: {total_thruput_rembang:.2f}")

    print(f"\nTotal Terminal Cost untuk BALONGAN: {total_terminal_cost_balongan:.2f}")
    print(f"Total Terminal Cost untuk CILACAP: {total_terminal_cost_cilacap:.2f}")

    print(f"\nTotal Supply Cost untuk OPSICO: {total_supply_cost_opsico:.2f}")
    print(f"Total Supply Cost untuk PEL: {total_supply_cost_pel:.2f}")
    print(f"Total Supply Cost untuk REMBANG: {total_supply_cost_rembang:.2f}")
    print(f"Total Supply Cost untuk BALONGAN: {total_supply_cost_balongan:.2f}")
    print(f"Total Supply Cost untuk CILACAP: {total_supply_cost_cilacap:.2f}")
    
    # Print totals
    print("\nSummary:")
    print(f"Total Distance: {total_distance}")
    print(f"Total Transport cost : {total_transport_cost:.2f}")
    print(f"Total Terminal Cost: {total_terminal_cost:.2f}")
    print(f"Total Thruput Fee: {total_thruput_fee:.2f}")
    print(f"Total Supply Cost: {total_supply_cost:.2f}")
    print(f"Total cost :  {total_cost}")
else:
    print("No optimal solution found. Please check your constraints and data.")
