from setuptools import setup, find_packages
import codecs
import os

VERSION = '0.0.1'
DESCRIPTION = 'Connecting Python with Aspen HYSYS'
LONG_DESCRIPTION = 'A package for automate the process simulation software Aspen HYSYS via Python.'

# Setting up
setup(
    name="HyPy",
    version=VERSION,
    author="Aron Somogyi",
    author_email="<arosomogyi@molgroup.info>",
    description=DESCRIPTION,
    packages=find_packages(),
    install_requires=['pandas','numpy','win32com.clinet'],
    keywords=['python', 'video', 'stream', 'video stream', 'camera stream', 'sockets'],
    classifiers=[
        "Development Status :: 1 - Planning",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ]
)
