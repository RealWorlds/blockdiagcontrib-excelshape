# -*- coding: utf-8 -*- 
from setuptools import setup, find_packages 
import os, sys

version = '0.1.0'
long_description = open("README.txt").read()

classifiers = [
    "Development Status :: 4 - Beta",
    "Intended Audience :: System Administrators",
    "License :: OSI Approved :: Python Software Foundation License",
    "Programming Language :: Python",
    "Topic :: Software Development",
    "Topic :: Software Development :: Documentation",
    "Topic :: Text Processing :: Markup",
]

setup(
     name='blockdiagcontrib-excelshape',
     version=version,
     description='imagedrawer plugin for blockdiag',
     long_description=long_description,
     classifiers=classifiers,
     keywords=['diagram','generator'],
     author='MIZUNO Hiroki',
     author_email='mzpppp at gmail.com',
     url='http://bitbucket.org/mzp/blockdiagcontrib-excelshape',
     license='Apache License 2.0',
     packages=find_packages(),
     package_data = {'': ['buildout.cfg']},
     namespace_packages=['blockdiagcontrib_excelshape'],
     include_package_data=True,
     install_requires=[
        'blockdiag>=1.2.1',
        'webcolors>=1.3',
        'setuptools'
     ],
     entry_points="""
        [blockdiag_imagedrawers]
        excelshape = blockdiagcontrib_excelshape.excelshape
     """,
)

