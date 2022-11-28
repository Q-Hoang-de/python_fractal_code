import sys
from ContractGenerator import ContractGenerator

employeeName = 'name'
employeeAddress = 'adreß'
beginDate = 'today'
title = 'hello'
salary = '1200,00'
salaryInWord = 'hello'
hourWage = '9,60'
terminationPeriod = 'auf unbestimmte Zeit'
probation = '01.06.2022'
contractor = ContractGenerator(
    employeeName,
    employeeAddress,
    beginDate,
    title,
    salary,
    salaryInWord,
    hourWage,
    terminationPeriod,
    probation)

contractor.generateContract()