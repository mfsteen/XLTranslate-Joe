# XLTranslate

## Installation

To avoid pain, install h5py using the OS's package manager. In the
case of Debian:

$ apt-get install python-h5py

Setup a virtual environment with access to site-packages:

$ virtualenv --system-site-packages venv
$ source venv/bin/activate

Install the xltranslate package in the virtual env.

$ python setup.py develop

This should generate the 'xltranslate' script in the venv/bin/ folder.
