from setuptools import setup, find_packages

setup(
    name="MyExcelAddin",
    version="1.0",
    packages=find_packages(),
    install_requires=[
        'xlwings>=0.30.0',
        'pywin32>=305'
    ],
) 