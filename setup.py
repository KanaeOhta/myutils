from setuptools import setup, find_packages

setup(
    name='myutils',
    version='0.0.1',
    author='Kanae Ohta',
    author_email='example@example.com',
    url='https://github.com/KanaeOhta/myutils.git',
    descriptions='python scripts',
    long_description='',
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7'
)