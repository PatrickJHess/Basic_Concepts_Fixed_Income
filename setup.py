from setuptools import setup

# Read the requirements.txt file so pip knows what to install
with open('requirements.txt') as f:
    required_packages = f.read().splitlines()

setup(
    name='basic_concepts_fixed_income',  # The name pip will use to register it
    version='0.1.0',
    py_modules=['module_ basic_concepts_fixed_income'],       # <--- Replace with your actual module's name (e.g., 'pricing_core')
    install_requires=[
        'numpy',
        'pandas<3.0',
        'openpyxl',
        'pathvalidate',
        'matplotlib',
        'python-dateutil',
        'holidays',
        'ipynbname',
        'beautifulsoup4',
        'pandas_market_calendars',
        'scipy',
        'pandas-datareader'
],
)
