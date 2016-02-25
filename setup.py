import os

from setuptools import setup, find_packages

here = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(here, 'README.md')) as f:
    README = f.read()
with open(os.path.join(here, 'CHANGES.txt')) as f:
    CHANGES = f.read()

requires = [
    'openpyxl',
    'h5py',
]

tests_require = requires + ['mock', ]

setup(name='xltranslate',
      version='0.1',
      description='xltranslate',
      long_description=README + '\n\n' + CHANGES,
      classifiers=[
        "Programming Language :: Python",
        ],
      author='Joe Steeve',
      author_email='js@hipro.co.in',
      url='',
      keywords='data_extraction',
      packages=find_packages(),
      include_package_data=True,
      zip_safe=False,
      install_requires=requires,
      tests_require=tests_require,
      extras_require={'test': tests_require},
      test_suite="xltranslate",
      entry_points="""\
      [console_scripts]
      xltranslate = xltranslate.script:main
      """,
)
