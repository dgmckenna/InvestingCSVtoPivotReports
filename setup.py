from cx_Freeze import setup,Executable

setup(
    name = 'Test',
    version = '0.1',
    description = 'Test Desc',
    executables = [Executable('InvestingCSVtoPivotReports.py')]
    )
    
