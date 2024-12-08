from ortools.linear_solver import pywraplp

solver = pywraplp.Solver.CreateSolver("GLOP")

inf = solver.infinity()
x = solver.IntVar(0, inf, "x")
y = solver.IntVar(0, inf, "y")

cons1 = solver.Add(x+2*y <= 20)
cons2 = solver.Add(3*x+y<=30)

solver.Maximize(3 * x + 4 * y)


status = solver.Solve()


if status == pywraplp.Solver.OPTIMAL:
    print(f'Solution:')
    print(f'x = {x.solution_value()}')
    print(f'y = {y.solution_value()}')
    print(f'Objective value = {solver.Objective().Value()}')
else:
    print('The problem does not have an optimal solution.')


