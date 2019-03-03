from setuptools import setup

setup(name='kfi_daily',
      version='1.1.1',
      description='equity tracking script',
      url='https://github.com/csb-comren/data_processing/',
      author='Khaleel Arfeen',
      author_email='k.a@unb.ca',
      license='MIT',
      packages=['kfi_daily'],
      install_requires=[
          'openpyxl',
          'requests',
      ],
      zip_safe=False)
