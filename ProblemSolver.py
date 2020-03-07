import ExcelImporter as EI
import pulp as p 

# 1. Read data from excel
products, orders = EI.GetData()

# 2. Define problem
LProblem = p.LpProblem('CasademontOptimization', p.LpMinimize)

# 3. Define problem parameters
time_horizon = p.LpVariable('time_horizon', 1, cat='Integer')

# Order params
delivery_day = p.LpVariable('delivery_day', 0, time_horizon, 'Integer')
total_mass = p.LpVariable('total_mass', 0, cat='Integer')
required_lean = p.LpVariable('required_lean', 0, time_horizon, cat='Integer')

# set of suppliers
# set of bodyparts

# MeatTypes = suppliers x bodyparts
# MeatTypes params
provided_lean = p.LpVariable('provided_lean', 0, cat='Integer')
max_order     = p.LpVariable('max_order', 0, cat='Integer')
min_order     = p.LpVariable('min_order', 0, cat='Integer')
price         = p.LpVariable('price', 0, cat='Integer')
initial_stock = p.LpVariable('initial_stock', 0, cat='Integer')


# 4. Set problem variables

######################
# DECISION VARIABLES #
######################

# amount of certain meattype boght on a given dat in kg
amount = p.LpVariable('amount', 0, cat='Integer')
# the amount we use from a certain meattype in the mix for an order in kg
uset = p.LpVariable('use', 0, cat='Integer')

########################
# CALCULATED VARIABLES #
########################

# value is 1 if we buy an amount o meattype on a given day
buy = p.LpVariable('buy', 0, 1, 'Binary')
# the amount of meattype that we stock for the next day
stock = p.LpVariable('stock', 0, cat='Integer')

###############
# CONSTRAINTS #
###############

# Initialize stock for the first day based on the given parameter 