import ExcelImporter as EI
import pulp as p 

# 1. Read data from excel
providers, orders = EI.GetData()

# 2. Define problem
LProblem = p.LpProblem('CasademontOptimization', p.LpMinimize)

# 3. Define problem variables
