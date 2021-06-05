#!/usr/bin/env python

"""from distutils.core import setup

#from os.path import abspath, join,pardir
#import src
#package_dir = abspath(join(__file__,pardir,"src"))


setup(name='pyPST2EML',
      version=src.__version__,
      description='Python tools for Outlook .PST to .eml translations',
      author='Matt-CHV',
      author_email='matthieu.chevrier.work@gmail.com',
      url='https://github.com/matt-chv',
      packages=['pyPST2EML'],
	  package_dir={'pyPST2EML': package_dir},
     )

"""

import setuptools
import src
from os.path import abspath, join, pardir

with open("README.md", "r") as fh:
    long_description = fh.read()
    
package_dir = abspath(join(__file__,pardir,"src"))
    
setuptools.setup(
    name="pyPST2EML",
    version=src.__version__,
    author="Matt-CHV",
    author_email="contact@matthieuchevrier.com",
    description="A small package to convert PST to folders of text searchable files with .eml (RFC-822)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/matt-chv/pyPST2EML",
    packages=['pyPST2EML'],
	 package_dir={'pyPST2EML': package_dir},
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Windows 10",
    ],
    install_requires=[
          'pywin32','python-dateutil'
      ],
    python_requires='>=3.7',
)