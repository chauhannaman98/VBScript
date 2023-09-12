' variables
dim salary
salary= 60000
dim tax_percent
tax_percent = 20
dim retirement
retirement = 5
dim health_plan
health_plan = 200

dim monthly_salary

monthly_salary = (salary - ((tax_percent/100)*salary)_
                 - ((retirement/100)*salary) - health_plan)/12
            
MsgBox "Adam receives $"&monthly_salary&" as monthly salary after all deductions!",0,"Techmirtz"