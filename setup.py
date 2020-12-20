from setuptools import setup, find_packages

setup(
    name='jsonexcel',
    version='1.0',
    author='Kanae Ohta',
    author_email='kanae5321@gmail.com',
    url='https://github.com/taKana671/JsonExcel.git',
    descriptions='Export JSON-format data to Excel file or data in Excel file to JSON file.',
    long_description='README',
    long_description_content_type="text/markdown",
    packages=find_packages(),
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3.8",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.8'
)
